#!/bin/bash

# 自己署名証明書を生成するスクリプト
# OpenSSLが必要です

CERT_DIR="./certs"
KEY_FILE="$CERT_DIR/server.key"
CERT_FILE="$CERT_DIR/server.crt"

echo "Generating self-signed certificate for development..."

# 秘密鍵を生成
openssl genrsa -out $KEY_FILE 2048

# 証明書を生成（365日有効）
openssl req -new -x509 -key $KEY_FILE -out $CERT_FILE -days 365 \
  -subj "/C=JP/ST=Tokyo/L=Tokyo/O=YourCompany/CN=localhost"

echo "Certificate generated successfully!"
echo "Key file: $KEY_FILE"
echo "Certificate file: $CERT_FILE"

# Windowsの場合はmkcertの使用を推奨
echo ""
echo "For Windows users, consider using mkcert:"
echo "1. Install mkcert: https://github.com/FiloSottile/mkcert"
echo "2. Run: mkcert -install"
echo "3. Run: mkcert -key-file certs/server.key -cert-file certs/server.crt localhost 127.0.0.1"
