# Command line PT-JPLsm docker image
# Use miniconda
FROM continuumio/miniconda3

# copy package content
COPY environment.yml .

# install libraries
RUN apt-get update && apt-get install libgl1 \
 libavcodec-dev libavformat-dev libswscale-dev \
 libgstreamer-plugins-base1.0-dev libgstreamer1.0-dev \
 libgtk2.0-dev libgtk-3-dev \
 libpng-dev libjpeg-dev libopenexr-dev libtiff-dev libwebp-dev -y

# recreate and activate the environment
RUN conda env create -f environment.yml
RUN echo "source activate transcribing_drc_data_environment" > ~/.bashrc
ENV PATH /opt/conda/envs/env/bin:$PATH
