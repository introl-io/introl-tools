#!/usr/bin/env bash

echo "Setting up dev certs..."

dotnet tool update -g linux-dev-certs
dotnet-linux-dev-certs install

sudo -E dotnet dev-certs https --export-path /usr/local/share/ca-certificates/dotnet-dev-cert.crt --format pem

sudo update-ca-certificates

echo "Dev certs updated successfully."