# jpmphotochem
Predictive simulation of photochemical experiments

The source code included here contains algorithms that are introduced in the publication "Predicting Wavelength-Dependent Photochemical Reactivity and Selectivity" 
by Jan Philipp Menzel, Benjamin B. Noble, James P. Blinco and Christopher Barner-Kowollik. 

Author: The code was created by Jan Philipp Menzel.
Contact by e-mail: j.p.menzel (at) gmail.com

Content of this README file: Brief description of each algorithm (12 - 32), System requirements (34 - 36), Installation guide (38 - 40), Demo (42 - 45), Instructions for use (47 - 49)

Description for jpmenzelqled.py:

Purpose of the algorithm: Quantitative prediction of wavelength, photon number, time and concentration dependent conversion of photoreaction employing an LED.
Notes: The source code is designed to predict conversion of thioether-substituted o-methylbenzaldehyde A with N-ethylmaleiminde NEM (refer to the above-mentioned publication) 
using LED 2 (emission centered around 343 nm) in the respective 3D-printed photoreactor. 
Detailed Notes: The algorithm requests manual input, imports uv/vis data from a respective excel file, makes lists for spatial distribution of reactands and products, 
calculates time dependent development of overall conversion for wavelengths of the respective LED at requested amount of reactands as well as calculates light attenuation maps.
Different LEDs can be used for predictions and the respective data can be entered in the source code.

Description for jpmenzelqselectivity.py:

Purpose of the algorithm: Quantitative prediction of the wavelength-dependent selectivity of two competing photoreactions using monochromatic light.
The algorithm requests manual input, imports uv/vis data from a respective excel file, makes lists for spatial distribution of reactands and products, 
calculates time dependent development of overall conversion for varied monochromatic wavelengths at requested amount of reactands as well as calculates the required photon count.

Description for jpmenzelqledorthogonal.py:

Purpose of the algorithm: Quantitative prediction of wavelength, photon number, time and concentration dependent conversion of photoreaction employing LEDs.
This algorithm simulates the competing photoreaction between Dodecyl-thioether o-methylbenzaldehyde and a diaryltetrazole (N-Phenyl-p-OMe / C-Phenyl-p-methylester).
The algorithm requests manual input, imports uv/vis data from a respective excel file, makes lists for spatial distribution of reactands and products, 
calculates time dependent development of overall conversion for wavelengths of the respective LED at requested amount of reactands as well as calculates light attenuation maps.

System requirements:

A standard python interpreter (e.g. Visual Studio Code) on a common system (e.g. desktop PC or laptop with windows, 8 GB RAM, 2.4 GHz) is expected  to be more than sufficient to run all scripts. The libraries 'math', openpyxl' and 'datetime' are used by the algorithms. Microsoft excel is required to modify *.xlsx files.

Installation guide:

To install Visual Studio Code, download from https://code.visualstudio.com/. Follow the installation instructions and create a workspace. Place into the folder associated to the workspace the three source code files 'jpmenzelqled.py', 'jpmenzelselectivity.py' and jpmenzelorthogonal.py' as well as the excel files 'jpmenzelqledoutputread.xlsx', 'jpmenzelqselectivityoutputread.xlsx', 'jpmenzeluvvisqled.xlsx' and 'jpmenzeluvvisqselectivity.xlsx'.

Demo:

In Visual Studio Code, in the explorer (Ctrl+Shift+E), click on the desired script. Before running any simulation, make sure that the appropriate excel files containing suitable data is in the correct location. 
To start a script, press CTRL+F5. In the terminal, the algorithm will display information and ask for input, before the iterative calculation starts. When the simulation is done, the name of the created file is displayed, which contains all relevant results from the simulation. The created excel file can be found in the folder associated with the workspace. Copy and rename the file to a location outside the workspace, if you wish to keep the results. Each file contains information about the user input as well as about the calculation results. If the results file is not copied and renamed, it may be overwritten, when the script is run again. 

Instructions for use:

To apply the predictive algorithm the other photochemical experiments, first ensure that all required parameters for making predictions are available, as discussed in the publication "Predicting Wavelength-Dependent Photochemical Reactivity and Selectivity" by Jan Philipp Menzel, Benjamin B. Noble, James P. Blinco and Christopher Barner-Kowollik, Nature Communications. In the folder associated to the workspace, make a copy of the relevant script and modify to embed the information relevant to your experiment. This may include measuring an emission spectrum of the light source, fitting the spectrum (e.g. with a sum of 6 gaussian functions) and entering the parameters of these into the source code as well as measuring quantum yields and entering these into the source code. UV Vis spectra need to be entered into respective excel files in analogy to the examples provided. If the data of the obtained or available UV Vis spectrum is not formatted as in the example spectrum, correct this (e.g. using excel and the VLOOKUP function).




