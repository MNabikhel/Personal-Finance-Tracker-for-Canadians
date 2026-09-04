FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive

# System packages the build and test suite depend on:
#   - python3 / python3-pip build the workbooks and run the suite
#   - libreoffice-calc + python3-uno recalculate both editions and run the VBA
#     in headless LibreOffice Basic, which the tests rely on as an independent
#     implementation of OOXML and the Basic dialect
# Repository-pinned Python dependencies are installed from requirements-dev.txt
# in the install phase, once the source has been checked out.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        ca-certificates \
        curl \
        git \
        python3 \
        python3-pip \
        libreoffice-calc \
        python3-uno \
    && rm -rf /var/lib/apt/lists/*

USER ubuntu
