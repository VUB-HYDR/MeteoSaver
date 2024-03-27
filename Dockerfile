# Command line PT-JPLsm docker image
# Use miniconda
FROM continuumio/miniconda3

# copy package content
COPY environment.yml .

# recreate and activate the environment
RUN conda env create -f environment.yml
RUN echo "source activate transcribing_drc_data_environment" > ~/.bashrc
ENV PATH /opt/conda/envs/env/bin:$PATH
