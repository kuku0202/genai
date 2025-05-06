PDF to PowerPoint(Lambda)
- First, the lambda function will read the incoming event from DynamoDB to extract the file_key and user, where the file_key is the S3 path of uploaded PDF in slidevox-pdf-storage 
- Download the PDF to /tmp/input.pdf using boto3
- Call process_pdf to extract text and figures using PyMuPDF and LayoutParser, generate slide scripts and teacher scripts via GPT and store each page's results in output/page_x folders
- Next, upload page assets to S3 including PDF text, teacher scripts, slide scripts, and relevant images
- Call create_pptx() to combine extracted content into a structured pptx and save to the tmp directory
- Upload .pptx to S3, including user_email 
- Send the PowerPoint to the user using AWS SES with a MIME-formatted email
- Send Task to SQS by creating a JSON message with ppt_key, pdf_name, user_email

PowerPoint to video(EC2)
- First poll the SQS Queue by calling the sqs.receive_message(). Once received, it extracts ppt_key, pdf_name and user email
- Create temporary directory under /tmp/project_[pdf_name]. Since each temporary directory is unique, it allow multiple task processing
- Download the .pptx file from S3 to /tmp/project_[pdf_name]/slides.pptx and extracted folders from pdf-to-ppt-output/output/pages/[pdf_name]/ into the same temp directory
- Call extract-ppt.py (LibreOffice backend) to convert the PowerPoint file into slide images (PNG format), saved in slides/
- Calls video.py to generate audio narration using TTS (e.g., gTTS or edge-tts) from teacher scripts, merge audio with slide images using FFmpeg and store video
- Upload the generated video to the ppt-to-video-output bucket
- Use generate_presigned_url() to create a time-limited S3 download link (valid for 24 hours)
- After video sent, cleanup the directory and delete that message
