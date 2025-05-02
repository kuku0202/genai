# import os
# import boto3
# import shutil
# import subprocess

# s3 = boto3.client('s3')

# def download_folder(bucket, prefix, local_dir):
#     paginator = s3.get_paginator('list_objects_v2')
#     for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
#         for obj in page.get('Contents', []):
#             key = obj['Key']
#             rel_path = key[len(prefix):]
#             if not rel_path: continue  # skip base folder
#             local_path = os.path.join(local_dir, rel_path)
#             os.makedirs(os.path.dirname(local_path), exist_ok=True)
#             s3.download_file(bucket, key, local_path)

# def upload_file(bucket, key, local_path):
#     s3.upload_file(local_path, bucket, key)
    
# def get_latest_prefix(bucket, base_prefix):
#     paginator = s3.get_paginator('list_objects_v2')
#     result = paginator.paginate(Bucket=bucket, Prefix=base_prefix, Delimiter='/')
#     latest_prefix = ""
#     latest_time = None
#     for page in result:
#         for cp in page.get('CommonPrefixes', []):
#             folder = cp['Prefix']
#             objs = s3.list_objects_v2(Bucket=bucket, Prefix=folder)
#             for obj in objs.get('Contents', []):
#                 if not latest_time or obj['LastModified'] > latest_time:
#                     latest_prefix = folder
#                     latest_time = obj['LastModified']
#     return latest_prefix


# def main():
#     input_bucket = "pdf-to-ppt-output"
#     output_bucket = "ppt-to-video-output"
#     base_prefix = "output/pages/"
#     tmp_dir = "/tmp/project"
#     slides_dir = os.path.join(tmp_dir, "slides")
#     scripts_dir = os.path.join(tmp_dir, "scripts")
#     output_video = os.path.join(tmp_dir, "presentation_video.mp4")

#     # Clean up temp
#     if os.path.exists(tmp_dir):
#         shutil.rmtree(tmp_dir)
#     os.makedirs(slides_dir)
#     os.makedirs(scripts_dir)

#     latest_prefix = get_latest_prefix(input_bucket, base_prefix)
#     print(f"📁 Latest PDF folder: {latest_prefix}")

#     download_folder(input_bucket, latest_prefix, tmp_dir)

#     # Extract PDF filename (like 'irag' from 'output/pages/irag/')
#     slide_dir_prefix = os.path.normpath(latest_prefix).split("/")[-1]
#     print(f"slide_dir_prefix: {slide_dir_prefix}")
#     pptx_key = f"output/{slide_dir_prefix}_presentation.pptx"
#     print(f"pptx_key to download: {pptx_key}")
#     pptx_path = os.path.join(tmp_dir, "slides.pptx")
#     s3.download_file(input_bucket, pptx_key, pptx_path)

#     # Extract slides using LibreOffice
#     subprocess.run(["python3", "extract-ppt.py", pptx_path, "--output_dir", slides_dir], check=True)

#     # Generate video
#     subprocess.run([
#         "python3", "video.py",
#         "--slides_dir", slides_dir,
#         "--scripts_dir", tmp_dir,
#         "--output_dir", tmp_dir,
#         "--output_video", output_video
#     ], check=True)

#     video_key = f"{slide_dir_prefix}_video.mp4"
#     upload_file(output_bucket, video_key, output_video)
#     print(f"✅ Uploaded video to s3://{output_bucket}/{video_key}")

# if __name__ == "__main__":
#     main()







import os
import boto3
import shutil
import subprocess
import json
import time
import base64
# import threading
from concurrent.futures import ThreadPoolExecutor

executor = ThreadPoolExecutor(max_workers=4)

s3 = boto3.client('s3', region_name = 'us-east-2')
sqs = boto3.client('sqs', region_name='us-east-2') 
QUEUE_URL = 'https://sqs.us-east-2.amazonaws.com/585768174219/ppt-generation-complete-v3'
INPUT_BUCKET = "pdf-to-ppt-output"
OUTPUT_BUCKET = "ppt-to-video-output"
# TMP_DIR = "/tmp/project"



def download_folder(bucket, prefix, local_dir):
    paginator = s3.get_paginator('list_objects_v2')
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get('Contents', []):
            key = obj['Key']
            rel_path = key[len(prefix):]
            if not rel_path: continue  # skip base folder
            local_path = os.path.join(local_dir, rel_path)
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            s3.download_file(bucket, key, local_path)

def upload_file(bucket, key, local_path):
    # s3.upload_file(local_path, bucket, key)
    s3.upload_file(
        local_path,
        bucket,
        key,
        ExtraArgs={
            "ContentType": "video/mp4"
        }
    )

    




import base64

ses_client = boto3.client("ses", region_name="us-east-2")
EMAIL_FROM = "info@theslidevox.com"


def generate_presigned_url(bucket_name, key, expiration=86400):
    return s3.generate_presigned_url(
        ClientMethod='get_object',
        Params={'Bucket': bucket_name, 'Key': key},
        ExpiresIn=expiration
    )

def send_email_with_link(download_url, recipient_email):
    try:
        print(f"📧 Sending download link to {recipient_email}")
        response = ses_client.send_email(
            Source=EMAIL_FROM,
            Destination={"ToAddresses": [recipient_email]},
            Message={
                "Subject": {"Data": "Your Presentation Video is Ready!"},
                "Body": {
                    "Text": {
                        "Data": f"""
Hi there,

Your presentation video is ready for download:

🔗 {download_url}

Important, the link above is only valid for 24 hours.
Thanks for using SlideVox!

- The SlideVox Team
"""
                    }
                }
            }
        )
        print(" Email sent:", response["MessageId"])
    except Exception as e:
        print(" Error sending email:", e)












def process_ppt(ppt_key, pdf_name, user_email, TMP_DIR):
    """Download PPT and scripts, extract slides, generate video, and upload."""
    slides_dir = os.path.join(TMP_DIR, "slides")
    output_video = os.path.join(TMP_DIR, f"{pdf_name}_video.mp4")
    pptx_path = os.path.join(TMP_DIR, "slides.pptx")
    scripts_prefix = f"output/pages/{pdf_name}/"

    # Clean workspace
    if os.path.exists(TMP_DIR):
        shutil.rmtree(TMP_DIR)
    os.makedirs(slides_dir, exist_ok=True)

    # Download files
    print(f"📥 Downloading PPTX from: {ppt_key}")
    s3.download_file(INPUT_BUCKET, ppt_key, pptx_path)

    print(f"📥 Downloading scripts from prefix: {scripts_prefix}")
    download_folder(INPUT_BUCKET, scripts_prefix, TMP_DIR)

    # Extract slides
    print("🖼️ Extracting slides using LibreOffice...")
    subprocess.run(["python3", "extract-ppt.py", pptx_path, "--output_dir", slides_dir], check=True)

    # Generate video
    print("🎞️ Generating video...")
    subprocess.run([
        "python3", "video.py",
        "--slides_dir", slides_dir,
        "--scripts_dir", TMP_DIR,
        "--output_dir", TMP_DIR,
        "--output_video", output_video
    ], check=True)

    # Upload video
    video_key = f"{pdf_name}_video.mp4"
    upload_file(OUTPUT_BUCKET, video_key, output_video)
    print(f"✅ Uploaded video to s3://{OUTPUT_BUCKET}/{video_key}")
    
    if user_email:
        # s3_url = f"https://{OUTPUT_BUCKET}.s3.amazonaws.com/{video_key}"
        s3_url = generate_presigned_url(OUTPUT_BUCKET, video_key)
        send_email_with_link(s3_url, user_email)


    if os.path.exists(TMP_DIR):
        shutil.rmtree(TMP_DIR)
    print(" Cleaned up /tmp/project after video upload.")

def poll_queue():
    """Continuously poll SQS queue for messages and process them."""
    print(" Polling SQS queue...")
    
    while True:
        # keep_latest_sqs_message(QUEUE_URL)
        
        messages = sqs.receive_message(
            QueueUrl=QUEUE_URL,
            MaxNumberOfMessages=1,
            WaitTimeSeconds=20
        )
        print("Video receive message:", messages)
        if "Messages" not in messages:
            time.sleep(10)
            continue

        for message in messages["Messages"]:
            receipt_handle = message["ReceiptHandle"]

            try:
                # if os.path.exists(TMP_DIR):
                #     shutil.rmtree(TMP_DIR)
                #     print("Clean up before processing")
                
                body = json.loads(message["Body"])
                ppt_key = body["ppt_key"]
                pdf_name = body["pdf_name"]
                user_email = body.get("user_email", "info@theslidevox.com")
                tmp_dir = f"/tmp/project_{pdf_name}"
                print(f" Received: ppt_key={ppt_key}, pdf_name={pdf_name}")

                # process_ppt(ppt_key, pdf_name, user_email, tmp_dir)
                # sqs.delete_message(QueueUrl=QUEUE_URL, ReceiptHandle=receipt_handle)
                # threading.Thread(target=process_and_cleanup, args=(ppt_key, pdf_name, user_email, tmp_dir, receipt_handle)).start()
                executor.submit(process_and_cleanup, ppt_key, pdf_name, user_email, tmp_dir, receipt_handle)
                # print(" Deleted message from queue.")
                
            except Exception as e:
                print(f" Error processing message: {e}")
            
            # finally:
            #     sqs.delete_message(QueueUrl=QUEUE_URL, ReceiptHandle=receipt_handle)
            #     print(" Deleted message from queue.")

def process_and_cleanup(ppt_key, pdf_name, user_email, tmp_dir, receipt_handle):
    """Wrap process_ppt and cleanup inside a separate thread."""
    try:
        process_ppt(ppt_key, pdf_name, user_email, tmp_dir)
        sqs.delete_message(QueueUrl=QUEUE_URL, ReceiptHandle=receipt_handle)
    except Exception as e:
        print(f" Error inside thread: {e}")

if __name__ == "__main__":
    poll_queue()
