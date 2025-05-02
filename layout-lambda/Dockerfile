FROM public.ecr.aws/lambda/python:3.10

# Install system-level dependencies
RUN yum install -y \
    git gcc gcc-c++ make cmake \
    libjpeg-devel zlib-devel libxml2-devel libxslt-devel \
    poppler-utils tesseract which
    # tar xz

# # Download and install static ffmpeg
# RUN curl -L https://johnvansickle.com/ffmpeg/releases/ffmpeg-release-amd64-static.tar.xz \
#     | tar -xJ --strip-components=1 -C /usr/local/bin --wildcards '*/ffmpeg' '*/ffprobe'

# # Set PATH to include ffmpeg binaries
# ENV PATH="/usr/local/bin:$PATH"     

# Install Python build tools and PyTorch + Detectron2
RUN pip install --upgrade pip && \
    pip install openai==0.28 && \
    pip install wheel ninja && \
    pip install torch==2.0.0 torchvision==0.15.1 torchaudio==2.0.1 \
      --index-url https://download.pytorch.org/whl/cpu && \
    PIP_NO_BUILD_ISOLATION=0 pip install 'git+https://github.com/facebookresearch/detectron2.git'

# Install PyTorch and Ultralytics

# RUN pip install --upgrade pip && \
#     pip install openai==0.28 && \
#     pip install wheel ninja

# RUN pip install torch==2.0.0 torchvision==0.15.1 --index-url https://download.pytorch.org/whl/cpu
# RUN pip install ultralytics

# Install remaining dependencies
# RUN pip install layoutparser pymupdf pdf2image pillow numpy==1.23.5 python-pptx lxml boto3 gtts tqdm
RUN pip install layoutparser pymupdf pdf2image pillow numpy==1.23.5 python-pptx lxml boto3

# Copy handler
COPY lambda_func.py .
COPY ppt.py .
# COPY extract-ppt.py .
# COPY video.py .
COPY model/newest.pth /models/layoutparser/publaynet/model_final_new.pth

# AWS Lambda entrypoint
# Install FFmpeg
   
CMD ["lambda_func.lambda_handler"]




