FROM python:3.9-slim

# Configure apt and install packages
RUN apt-get update && \
    DEBIAN_FRONTEND=noninteractive apt-get install -y \
    gdal-bin \
    libgdal-dev \
    gcc \
    g++ \
    python3-dev \
    krb5-user \
    libkrb5-dev \
    libssl-dev \
    git \
    pkg-config \
    && rm -rf /var/lib/apt/lists/*

# Configure krb5-config non-interactively
ENV DEBIAN_FRONTEND=noninteractive
RUN echo "krb5-config krb5-config/default_realm string EXAMPLE.COM" | debconf-set-selections && \
    echo "krb5-config krb5-config/kerberos_servers string localhost" | debconf-set-selections && \
    echo "krb5-config krb5-config/admin_server string localhost" | debconf-set-selections && \
    apt-get update && \
    apt-get install -y krb5-config && \
    rm -rf /var/lib/apt/lists/*

# Set GDAL environment variables
ENV GDAL_VERSION=3.4.1
ENV CPLUS_INCLUDE_PATH=/usr/include/gdal
ENV C_INCLUDE_PATH=/usr/include/gdal

# Set working directory
WORKDIR /app

# Copy requirements file
COPY requirements.txt ./

# Install Python packages
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir wheel setuptools && \
    pip install --no-cache-dir -r requirements.txt

# Install gunicorn
RUN pip install gunicorn

# Copy application code
COPY . .

# Create log directory
RUN mkdir -p /app/logs && \
    chmod -R 777 /app/logs

# Set timezone
ENV TZ=America/Denver
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

# Expose port 8080
EXPOSE 8080

# Command to run the Flask app with gunicorn
CMD exec gunicorn --bind :8080 --workers 1 --threads 8 --timeout 0 main:app