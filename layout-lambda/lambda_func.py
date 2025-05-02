# import boto3
# import base64
# import os
# import shutil
# import mimetypes
# # import subprocess
# # import asyncio
# from ppt import process_pdf, create_pptx  
# # from video import VideoGenerator

# s3_client = boto3.client('s3')

# ses_client = boto3.client('ses', region_name='us-east-2')  # adjust the region as needed
# # Configuration constants
# OUTPUT_BUCKET = 'pdf-to-ppt-output'
# EMAIL_FROM = "info@theslidevox.com"
# def send_email_with_attachment(filename, recipient_email):
#     """Sends an email with the processed PPT attachment from S3 using AWS SES."""
#     try:
#         print(f": Fetching {filename} from S3 for email attachment...")
#         file_obj = s3_client.get_object(Bucket=OUTPUT_BUCKET, Key=filename)
#         file_content = file_obj["Body"].read()
#         attachment = base64.b64encode(file_content).decode()
#         # Appropriate content type for PPTX files
#         content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
#         print(f": Sending email with attachment: {filename} to {recipient_email}...")
#         response = ses_client.send_raw_email(
#     Source=EMAIL_FROM,
#     Destinations=[recipient_email],
#     RawMessage={
#         "Data": f"""From: {EMAIL_FROM}
# To: {recipient_email}
# Subject: Your Converted PPT is Ready
# MIME-Version: 1.0
# Content-Type: multipart/mixed; boundary="NextPart"

# --NextPart
# Content-Type: text/plain; charset="utf-8"
# Content-Transfer-Encoding: 7bit

# Hello,
# Please find attached your converted PPT file.

# Best regards,
# The SlideVox Team

# --NextPart
# Content-Type: {content_type}
# Content-Disposition: attachment; filename="{filename}"
# Content-Transfer-Encoding: base64

# {attachment}
# --NextPart--
# """
#     }
# )

#         print(" Email sent successfully:", response)
#     except Exception as e:
#         print(":x: Error sending email:", str(e))



# def lambda_handler(event, context):
    

    
#     bucket = event['Records'][0]['s3']['bucket']['name']
#     key = event['Records'][0]['s3']['object']['key']
    
#     pdf_local_path = '/tmp/input.pdf'
#     output_dir = '/tmp/output'
#     pptx_local_path = '/tmp/generated_presentation.pptx'
#     # slides_dir = '/tmp/slides'
#     # video_path = '/tmp/presentation_video.mp4'
    
#     s3_client.download_file(bucket, key, pdf_local_path)
#     if os.path.exists(output_dir):
#         shutil.rmtree(output_dir)
    
#     os.makedirs(output_dir, exist_ok=True)

#     original_obj = s3_client.head_object(Bucket=bucket, Key=key)
#     # print("Object metadata:", original_obj.get("Metadata", {}))
#     ######################################
    
    
#     # user_email = original_obj.get("Metadata",{}).get('user-email')
#     # if not user_email:
#     #     user_email = "unknown@example.com" 
#     # print(user_email)
#     metadata = original_obj.get("Metadata", {})
#     print("Metadata:", metadata)
#     user_email = metadata.get('user_email', "theslidevox@gmail.com")
#     print(user_email)
#     # Process PDF and create PPTX
#     process_pdf(pdf_local_path, output_dir)
    
#     for page_dir in os.listdir(output_dir):
#             full_page_path = os.path.join(output_dir, page_dir)
#             if os.path.isdir(full_page_path):
#                 for file_name in os.listdir(full_page_path):
#                     if file_name.endswith((".png", ".jpg", ".txt")):
#                         file_path = os.path.join(full_page_path, file_name)
#                         s3_key = f"output/pages/{os.path.splitext(os.path.basename(key))[0]}/{page_dir}/{file_name}"
#                         print(f"Uploading {file_name} to {s3_key}...")
#                         with open(file_path, "rb") as f:
#                             s3_client.upload_fileobj(
#                                 Fileobj=f,
#                                 Bucket=OUTPUT_BUCKET,
#                                 Key=s3_key
#                             )
                    
#     create_pptx(output_dir, pptx_filename=pptx_local_path)
    
#     ## Extract
#     # os.makedirs(slides_dir, exist_ok=True)
#     # subprocess.run(["python3", "extract-ppt.py", pptx_local_path, "--output_dir", slides_dir], check=True)

#     # generator = VideoGenerator(
#     #     slides_dir=slides_dir,
#     #     scripts_dir=output_dir,
#     #     output_dir="/tmp",
#     #     output_video="presentation_video.mp4",
#     #     tts_engine="gtts",
#     #     voice="en",
#     # )
#     # asyncio.run(generator.generate_video())
    
    
#     pptx_s3_key = f"output/{os.path.splitext(os.path.basename(key))[0]}_presentation.pptx"
    
#     # video_s3_key = f"output/{os.path.splitext(os.path.basename(key))[0]}_presentation.mp4"
    

#     content_type = mimetypes.guess_type(pptx_s3_key)[0] or "application/vnd.openxmlformats-officedocument.presentationml.presentation"
#     with open(pptx_local_path, "rb") as f:
#         s3_client.upload_fileobj(
#             Fileobj=f,
#             Bucket=OUTPUT_BUCKET,
#             Key=pptx_s3_key,
#             ExtraArgs={
#                 "Metadata": {
#                     "user_email": user_email
#                 },
#                 "ContentType": content_type
#             }
#         )
    
#     # with open(video_path, "rb") as f:
#     #     s3_client.upload_fileobj(f, OUTPUT_BUCKET, video_s3_key, ExtraArgs={
#     #         "ContentType": "video/mp4"
#     #     })
#     # Return the S3 path to the generated PPTX
#     pptx_s3_url = f"s3://{OUTPUT_BUCKET}/{pptx_s3_key}"
#     send_email_with_attachment(pptx_s3_key, user_email)
    
#     return {
#         'statusCode': 200,
#         'body': f"Presentation created and saved to {pptx_s3_url}"
#     }











import boto3
import base64
import os
import shutil
import mimetypes
import json
# import subprocess
# import asyncio
from ppt import process_pdf, create_pptx  
# from video import VideoGenerator

s3_client = boto3.client('s3')
ses_client = boto3.client('ses', region_name='us-east-2')  # adjust the region as needed

sqs = boto3.client('sqs', region_name='us-east-2')
QUEUE_URL = 'https://sqs.us-east-2.amazonaws.com/585768174219/ppt-generation-complete-v3'
# Configuration constants
OUTPUT_BUCKET = 'pdf-to-ppt-output'
EMAIL_FROM = "info@theslidevox.com"
def send_email_with_attachment(filename, recipient_email):
    """Sends an email with the processed PPT attachment from S3 using AWS SES."""
    try:
        print(f": Fetching {filename} from S3 for email attachment...")
        file_obj = s3_client.get_object(Bucket=OUTPUT_BUCKET, Key=filename)
        file_content = file_obj["Body"].read()
        attachment = base64.b64encode(file_content).decode()
        # Appropriate content type for PPTX files
        content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        print(f": Sending email with attachment: {filename} to {recipient_email}...")
        response = ses_client.send_raw_email(
    Source=EMAIL_FROM,
    Destinations=[recipient_email],
    RawMessage={
        "Data": f"""From: {EMAIL_FROM}
To: {recipient_email}
Subject: Your Converted PPT is Ready
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="NextPart"

--NextPart
Content-Type: text/plain; charset="utf-8"
Content-Transfer-Encoding: 7bit

Hello,
Please find attached your converted PPT file.

Best regards,
The SlideVox Team

--NextPart
Content-Type: {content_type}
Content-Disposition: attachment; filename="{filename}"
Content-Transfer-Encoding: base64

{attachment}
--NextPart--
"""
    }
)

        print(" Email sent successfully:", response)
    except Exception as e:
        print(":x: Error sending email:", str(e))



def lambda_handler(event, context):
    record = event['Records'][0]['dynamodb']['NewImage']
    key = record['file_key']['S']
    user_email = record['email']['S']
    bucket = "slidevox-pdf-storage"
    pdf_local_path = '/tmp/input.pdf'
    output_dir = '/tmp/output'
    pptx_local_path = '/tmp/generated_presentation.pptx'
    s3_client.download_file(bucket, key, pdf_local_path)
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    process_pdf(pdf_local_path, output_dir)
    
    for page_dir in os.listdir(output_dir):
            full_page_path = os.path.join(output_dir, page_dir)
            if os.path.isdir(full_page_path):
                for file_name in os.listdir(full_page_path):
                    if file_name.endswith((".png", ".jpg", ".txt")):
                        file_path = os.path.join(full_page_path, file_name)
                        s3_key = f"output/pages/{os.path.splitext(os.path.basename(key))[0]}/{page_dir}/{file_name}"
                        print(f"Uploading {file_name} to {s3_key}...")
                        with open(file_path, "rb") as f:
                            s3_client.upload_fileobj(
                                Fileobj=f,
                                Bucket=OUTPUT_BUCKET,
                                Key=s3_key
                            )
                    
    create_pptx(output_dir, pptx_filename=pptx_local_path)
    
    ## Extract
    # os.makedirs(slides_dir, exist_ok=True)
    # subprocess.run(["python3", "extract-ppt.py", pptx_local_path, "--output_dir", slides_dir], check=True)

    # generator = VideoGenerator(
    #     slides_dir=slides_dir,
    #     scripts_dir=output_dir,
    #     output_dir="/tmp",
    #     output_video="presentation_video.mp4",
    #     tts_engine="gtts",
    #     voice="en",
    # )
    # asyncio.run(generator.generate_video())
    
    
    pptx_s3_key = f"output/{os.path.splitext(os.path.basename(key))[0]}_presentation.pptx"
    
    # video_s3_key = f"output/{os.path.splitext(os.path.basename(key))[0]}_presentation.mp4"
    

    content_type = mimetypes.guess_type(pptx_s3_key)[0] or "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    with open(pptx_local_path, "rb") as f:
        s3_client.upload_fileobj(
            Fileobj=f,
            Bucket=OUTPUT_BUCKET,
            Key=pptx_s3_key,
            ExtraArgs={
                "Metadata": {
                    "user_email": user_email
                },
                "ContentType": content_type
            }
        )
    
    # with open(video_path, "rb") as f:
    #     s3_client.upload_fileobj(f, OUTPUT_BUCKET, video_s3_key, ExtraArgs={
    #         "ContentType": "video/mp4"
    #     })
    # Return the S3 path to the generated PPTX
    pptx_s3_url = f"s3://{OUTPUT_BUCKET}/{pptx_s3_key}"
    send_email_with_attachment(pptx_s3_key, user_email)
    
    
    
    
    #############################################################################
    
    message_body = {
        "ppt_key": pptx_s3_key,
        "pdf_name": os.path.splitext(os.path.basename(key))[0],
        "user_email": user_email
    }

    # sqs.send_message(
    #     QueueUrl=QUEUE_URL,
    #     MessageBody=json.dumps(message_body)
    # )
    
    response = sqs.send_message(
        QueueUrl=QUEUE_URL,
        MessageBody=json.dumps(message_body)
    )
    print(f" SQS message sent: MessageId={response['MessageId']}, Body={json.dumps(message_body)}")
    # print(json.dumps(response, indent=2))
    
        
    
    return {
        'statusCode': 200,
        'body': f"Presentation created and saved to {pptx_s3_url}"
    }