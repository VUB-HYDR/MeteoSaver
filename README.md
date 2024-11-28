# MeteoSaver v1.0

Here we present MeteoSaver v1.0, a machine-learning based software for the transcription of historical weather data.

## Directory structure

Below is the structure for this project.

```
├── README.md                                               <- This includes general information on the project
|                                                              and introduction of project structure and modules
|
├── OCR_HTR_models                                          <- Trained and serialized models, model predictions, or model summaries.
|                                                         
├── data
│   ├── 00_post1960_DRC_hydroclimate_datasheet_images       <- Sample images of climate data sheets from the INERA
|                                                              Yangambi archives, DRC, arranged in folders (representing station numbers)
|   ├── 01_metadata_INERA_stations                          <- Metadata for the INERA meteorological stations
│
├── docs                                                    <- Flow charts and keys for MeteoSaver included in Muheki et al. 2025.
|
├── results                                                 <- All results obtained from exceution of MeteoSaver v1.0 on the sample ten datasheets
|   |                                                          included in this repository as a Minimal Working Example                                                                                         
│   ├──01_pre_QA_QC_transcribed_hydroclimate_data           <- Original automatic transcription using MeteoSaver before QA/QC checks
|   |                                       
│   ├── 02_post_QA_QC_transcribed_hydroclimate_data         <- Transcription after using MeteoSaver QA/QC checks
|   |                                      
│   ├── 03_validation_transcibed_data                       <- Validation of the transcription using MeteoSaver (in comparison to manual transcription)
│   │
│   ├── 04_final_refined_daily_hydroclimate_data            <- Final refined (formatted) trancribed data, in .xlsx and .tsv (SEF) ready for upload
|   │
│   ├── 05_transient_transcription_output                   <- Transient outputs (during processing)
|
│   ├── 06_manually_transcribed_hydroclimate_data           <- Manually transcribed data (For validation purposes)
|
|
|                                       
├── src                                                     <- Modules (2-6): Transcribing code/scripts for MeteoSaver v1.0
│   ├── main.py                                             <- Main script to run all the modules 1-6 of MetoSaver (scripts)
|   |                                                          i.e. in order (i) configuration, (iI) image-preprocessing module, (iii) table and cell
|   |                                                          detection model, (iv) transcription, (v) quality assessment and control,
|   |                                                          and (vi) data formatting and upload
│   │
│   ├── image_preprocessing_module.py                       <- Script to carry out image preprocessing of the original scans
|   |                                                          of climate data records
│   │
│   ├── table_and_cell_detection_model.py                   <- Script to detect the table and cells from the already
|   |                                                          pre-processed images
│   │
│   ├── transcription.py                                    <- Script to detect the text within the detected cells using
|   |                                                          an Optical Character Recognition (OCR) or Handwritten Text
|   |                                                          Recognition (HTR) model of your choice.
|   │
│   ├── quality_assessment_and_quality_control.py           <- Script to perform QA/QC checks on the original automatically transcribed data
|   |                                       
│   ├── validation.py                                       <- Script to generates a visual comparison of daily maximum, minimum,
|   |                                                          and average temperatures between manually transcribed data and
|   |                                                          QA/QC checked transcribed data for a specific station
│   └── data_formatting_and_upload.py                       <- Script to select the confirmed data (from the QA/QC) and convert it both an excel file 
|                                                              and to the Station Exchange Format, as well plot timeseries per station
|   
│
├── Dockerfile                                              <- Docker install routine for a virtual environment
|
├──LICENSE                                                  <- Licence
|
├── configuration.ini                                       <- Module 1: Configuration. User-defined settings to ensure smooth running of MeteoSaver
|
├── environment.yml                                         <- The requirements file for reproducing the analysis environment
|                                                              Generated with `conda env export > environment.yml`
|
├── job_script.sh                                           <- Job script for HPC infrastructure users to run the software
|
└── setup.py                                                <- Make this project pip installable with `pip install -e` 

```

## Setup

Two ways of creating reproducible environments are provided, the general Conda environment and an isolated Docker environment based on a Conda base image.

> [!WARNING]
> It is adviced to work in isolated Docker environments in order to ensure reproducibility, future online deployments, but first and foremost security of your computer system. Pip and to a lesser degree Conda and their python environments are a known malware vector. Although the framework we present vets the loaded library we can not assure the safety of all dependencies created downstream. The use of the local non-containerized setup is therefore not recommended.

### Conda

To create an environment which is consistent use the environment file after installing Miniconda.

```bash
conda env create -f environment.yml
```

Activate the working environment using:

```bash
conda activate transcribing_drc_data_environment
```

### Docker

The dockerfile included provides a Conda environment ([see here for docker install instructions](https://docs.docker.com/engine/install/)).
You can build this docker image using the below command. This will download all required
python components and packages, while safeguarding (sandboxing) your system
from `pip` based [security issues](https://www.bleepingcomputer.com/news/security/pypi-suspends-new-user-registration-to-block-malware-campaign/). Once build locally no further downloads 
will be required.

```
# In the main project directory run
docker build -f Dockerfile -t transcribing_drc_data_environment .
```

To spin up a docker image using:

```
docker run -it -v /local_data:/docker_data_dir transcribing_drc_data_environment
```

## Modules
The figure below represents the modules in MeteoSaver v1.0

  ![Schematic representation of the modules in MeteoSaver v1.0](https://github.com/VUB-HYDR/MeteoSaver/blob/6a5238af498088d58173940e04fe8e5cf66567be/docs/Schematic%20representation%20of%20the%20modules%20in%20MeteoSaver%20v1.0.png)


## How to run MeteoSaver v1.0 
After setting up the python environment using the [environment.yml](https://github.com/VUB-HYDR/MeteoSaver/blob/b8138fa5a23f4ce40603cae8defd82d10734fdbd/environment.yml) file available on this repository, input the following settings in the [configuration module](https://github.com/VUB-HYDR/MeteoSaver/blob/b8138fa5a23f4ce40603cae8defd82d10734fdbd/configuration.ini) specific to your case study (sheets) before running:
1. General: Here, you specify the environment in which the scripts will run, i.e. ```local``` (Sequential processing on a personal computer) or ```hpc``` (Parallel processing using multiple processors, suitable for High Performance Computing (HPC) environments). This is set to ```local``` by default
2. Directories. Here, you specify the directories for the following: (i) all historical weather data sheet images in folders per station, (ii) pre-QA/QC transcribed data, (iii) post-QA/QC transcribed data, (iv) the final refined daily hydroclimate data (after all quality checks), (v) transient transcription output during processing, (vi) manually transcribed data (used for validation), (vii) alidation results comparing manually transcribed and the MeteoSaver transcribed data, and (viii) all the stations metadata.
3. Table and Cell Detection: User specifications for table and cell detection.
4. Transcription: User specifications related to the Optical Character Recognition/Handwritten Text Recognition (OCR/HTR).
5. QA/QC: Here, you specify parts of the transcribed table on which to perform QA/QC checks.
6. Data Formatting: Here, you specify the location of the date information in the tables, used for formatting the transcribed data to time series in .xlsx and .tsv (Station Exchange Format).

After inputting the configuration settings specific to your case study (see Table below), you can then run the [main.py](https://github.com/VUB-HYDR/MeteoSaver/blob/7aeab0f526b44056c062407df7cfe467e20a67d8/src/main.py) script which runs all the modules 1-6 of MeteoSaver i.e. in order (i) configuration, (iI) image-preprocessing module, (iii) table and cell detection model, (iv) transcription, (v) quality assessment and control, and (vi) data formatting and upload, and return results in the specified directories. 

### Minimal Working Example (MWE)
You can run the entire script in this repository as a Minimal Working Example (MWE) without modifying any configuration settings. Simply set up the Python environment and execute the script using [main.py](https://github.com/VUB-HYDR/MeteoSaver/blob/7aeab0f526b44056c062407df7cfe467e20a67d8/src/main.py).


### Configuration user-settings
The figure below describes all the configuration user-settings.
![Configuration_user_settings](https://github.com/VUB-HYDR/MeteoSaver/blob/4ddd56d52b3dda19afc6227595eba0d6ca843c30/docs/Configuration%20user%20settings.png)




## Authors
Derrick Muheki

Koen Hufkens

Bas Vercruysse

Krishna Kumar Thirukokaranam Chandrasekar

Wim Thiery


## Acknowledgements

This template was inspired by the [python_proj_template](https://github.com/pepaaran/python_proj_template).
