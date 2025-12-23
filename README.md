# NMAA_spectral_scanner
Programs for the spectral imaging of works of art using instrumentation connected to a motorized scanner.

# National Museum of Asian Art Spectral Scanning Software
Conservation and Scientific Research, National Museum of Asian Art, Smithsonian Institution, Washington, DC, United States of America
Matthew Clarke, clarkem@si.edu

# General Overview
These programs are for connection of an instrument to a motorized scanning system, with the with the purpose of spectral imaging of artwork. The examples below have been used for the connection of a reflectance imaging camera and an x-ray fluorescence instrument to a scanner. While this program was developed for specific instruments, it may be adapted for other purposes. The Visual Basic program was written to operate in VB .NET 8.0. For specific instrument commands, please see the manufacturer’s documentation.
Simple communication connection changes may be done through the COMSettings.ini in the bin folder. Other changes require opening the VB solution in a program such as Visual Studio.
(An article describing the application of this software for the analysis of Asian artwork has been submitted. This document will be updated, pending acceptance/publication.)

# Spectral Imaging with VNIR camera
VB Program: SOC-VXM-VRO-v3-position
Equipment used and purpose:
Surface Optics 710-sCMOS VNIR (spectral camera)
Velmex motorized rails (motion)
Velmex encoders (position)
LabJack DAQ (trigger for camera)

Basic operation
1.	Connect to Velmex motorized rails and encoders
2.	Connect to LabJack
3.	Separately connect to Surface Optics software HyperScannerV2. The SOC 710 is only controlled through HyperScannerV2.
4.	Enter the number of lines to scan (iterations). The maximum for the SOC 710 is 1024.
5.	Enter the integration time of the camera. Ensure this matches the HyperScannerV2 software.
6.	Enter the field of view (FOV). This should be measured by test scans.
7.	Press the calculate button. The rate of motion will be determined by the integration time and the FOV.
8.	In the HyperScannerV2, the camera should be set to an external trigger and the number of frames collected should match the VB form value for iterations/lines.
9.	Press “Collect” in HyperScannerV2.
10.	Press “Start Scan” in the VB form
11.	Once data is collected, save in the HyperScannerV2 software.
Additional features
•	A movement box allows for simple positioning of the scanner.
•	The 10% button will move the system 10% of the FOV in that direction.
•	The system is designed to scan along the y-axis during acquisition.
•	Multiple scans can be overlapped using the movement buttons of the x and y axes.
•	A y-edge can be set to allow a return to a starting position.

# Spectral Imaging with XRF system
VB Program: vbDP5_VXM_VRO_Net8-32bit-V02-positionsave
Equipment used and purpose
Amptek SDD DP5 (x-ray fluorescence detection)
Velmex motorized rails (motion)
Velmex encoders (position)
DSD Tech (interlock control)

This VB program draws from Amptek XRF SDK software (https://www.amptek.com/software/dp5-digital-pulse-processor-software, downloaded 2022).  Please see the Amptek DP5SDK END-USER SOFTWARE LICENSE AGREEMENT (DP5SDK_EULA.pdf) for details of its use. The Amptek code has been updated here to VB .NET 8.0 for stability and long-term use.  
Specific configuration of the DP5 detector should be done in the Amptek DppMCA program. The VB software will use the last configuration file loaded to the detector.
This program does not control the x-ray tube, but it may be used to shut off the x-rays by means of an interlock. Always confirm x-rays are off before approaching the instrument. Follow local safety guidelines when using x-rays.
The system is designed to first scan along the y-axis.

Basic operation for the scanning mode
1.	Connect to the DP5 SSD.
2.	Connect to Velmex motorized rails and encoders
3.	Connect to DSD interlock control
4.	Set the project folder and name in the “Scan Name, Folder” button.
5.	Position the system of the top left corner. Click “Set Top Corner.”
6.	Position the system on the bottom right corner. Click “Set Bottom Corner.”
7.	Set the desired spacing between point in the “mm spacing” box
8.	Set the acquisition time (as real time) in the “ms real time”
9.	Press the “Calculate” button, which will estimate the time for the scan and the number of points.
10.	[Turn on x-rays in another program]
11.	Press “Start Scan”
12.	At the end of each y-scan, the time until completion will be updated.
13.	At the end of the full scan, the scanner will stop, and the interlock will toggle to shut off the x-rays.
14.	MCA files are saved at each data point. These can be combined into an hdf5 file by use of a python routine.
Additional features and notes
•	A file with the x-axis positions will be saved.
•	This VB program operates in 32-bit mode.

# Python Processing Files
Two processing files written in Python Jupyter Notebook are included. 
NMAA_XRF_Create_HDF5_2025.ipynb
This will take a set of MCA files and convert it to a 2D HDF5 file. This is x,y,spectrum format. It will not save other information (e.g., detector settings, live counts, etc.) for each point.
NMAA_VNIR_Process_Cubes_2025.ipynb
The second is used for batch processing of the VNIR image, such as flat-fielding and calibration of the images.
