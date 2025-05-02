# Use Debian-based image for better package support
FROM python:3.10-slim

ENV DEBIAN_FRONTEND=noninteractive
# System dependencies

RUN apt-get update && apt-get install -y libreoffice \
    curl ffmpeg xz-utils tesseract-ocr poppler-utils \
    git gcc g++ make cmake \
    libjpeg-dev zlib1g-dev libxml2-dev libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
RUN pip install --upgrade pip && \
    pip install openai==0.28 wheel tqdm gtts edge-tts boto3 python-pptx pdf2image pillow numpy lxml

# Copy all necessary scripts
COPY main.py .
COPY extract-ppt.py .
COPY video.py .

# Default command to run your main script
CMD ["python", "main.py"]