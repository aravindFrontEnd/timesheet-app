FROM registry.access.redhat.com/ubi8/python-39:latest

WORKDIR /app

# Install Tesseract OCR and dependencies as root
USER root
RUN yum update -y && \
    yum install -y tesseract tesseract-langpack-eng tesseract-devel && \
    yum clean all

# Switch back to non-root user for security
USER 1001

# Copy requirements and install Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Set environment variables
ENV PORT=8080
ENV FLASK_ENV=production
ENV TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata

EXPOSE 8080

# Run the application
CMD ["python", "app.py"]