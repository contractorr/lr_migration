# Data Migration
## Introduction
The code in this repository was used to migrate fund operation data from hundreds of Excel files into an alternative investment software suite called eFront. This source data was cleaned, transformed and output into an eFront Excel template so that it can be loaded into that software. Overall, more than 25,000 fund operations were imported. 

![image](https://user-images.githubusercontent.com/92265905/145563243-103bbe87-0f09-45c1-bf9d-76c1ad224f38.png)

The fund operation amounts were often classified using a text field in the source Excel files. This repository also contains code used to aggregate the data so that it can be analysed internally or reviewed and classified by other stakeholders. 

We did not require this code to be maintained or reused after the migration was complete. The emphasis was on speed rather than maintainability. 

## Key Files
There are lots of different files in this repository. The most relevant files are:
- mig_functions.py: this contains the majority of the code. It contains functions that are used by the other files. There will be docstrings to explain what each function does. 
- investee_fund_ops.py: this file collates all source data files and calls the relevant functions to migrate data for investee funds 
- managed_fund_ops.py: this file collates all source data files and calls the relevant functions to migrate data for managed funds
