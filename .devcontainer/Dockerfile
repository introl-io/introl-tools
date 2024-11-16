FROM mcr.microsoft.com/devcontainers/dotnet:9.0

# Install zsh
RUN apt-get update && apt-get install -y zsh && apt-get clean
# Change the default shell to zsh
RUN chsh -s /bin/zsh