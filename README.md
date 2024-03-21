# Data_Rescue_Congo_DRC
Here we undertake data transcription of millions of daily precipitation and temperature records collected within the Congo Basin.

## Setting up the repository and environment


## Directory structure
Below is the structure for this project.
```
├── README.md                            <- This includes general information on the project and introduction of project structure and modules
├── LICENSE
|
├── data
│   ├── 10_sample_different_images       <- Sample images of climate data sheets from the INERA Yangambi archives, DRC.
│  
├── notebook                             <- Jupyter notebook for exploration of the code only, as well examples of outputs from the jupyter notebooks such as Ms Excel with OCR results and clipped images.
│
├── transcribing_model                   <- Trancribing code/scripts for use in this project.
│   ├── main.py                          <- Script to run all the different modules (scripts) i.e. in order (1) image preprocessing module, (2) table detection model, and (3) transcription model
│   │
│   ├── image_preprocessing.py           <- Scripts to download or generate data.
│   │   └── 01_clean_dataset.py
│   │
│   ├── features       <- Scripts to turn raw data into features for modeling.
│   │   └── 02_build_features.py
│   │
│   ├── models         <- Scripts to train models and then use trained models to make
│   │   │                 predictions.
│   │   ├── 03_train_model.py
│   │   └── 04_predict_model.py
│   │
│   └── visualization  <- Scripts to create exploratory and results oriented visualizations
│       └── 05_visualize.py
│
├── references         <- Data dictionaries, manuals, bibliography (.bib)
│
├── reports            <- Generated analysis as HTML, PDF, LaTeX, etc.
│   └── figures        <- Generated graphics and figures to be used in reporting
│
├── environment.yml    <- The requirements file for reproducing the analysis environment, e.g.
│                         generated with `conda env export > environment.yml`
│
├── setup.py           <- Make this project pip installable with `pip install -e`
|
├── src                <- Source code for use in this project.
│   ├── __init__.py    <- Makes src a Python module.
│   │
│   ├── data           <- Scripts to download or generate data.
│   │   └── 01_clean_dataset.py
│   │
│   ├── features       <- Scripts to turn raw data into features for modeling.
│   │   └── 02_build_features.py
│   │
│   ├── models         <- Scripts to train models and then use trained models to make
│   │   │                 predictions.
│   │   ├── 03_train_model.py
│   │   └── 04_predict_model.py
│   │
│   └── visualization  <- Scripts to create exploratory and results oriented visualizations
│       └── 05_visualize.py
│
└── .gitignore         <- Indicates which files should be ignored when pushing.
```



This template was inspired by the [python_proj_template](https://github.com/pepaaran/python_proj_template).


