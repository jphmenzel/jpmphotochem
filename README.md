# jpmphotochem
Predictive simulation of photochemical experiments
The source code included here contains algorithms that are introduced in the publication "Predicting Wavelength-Dependent Photochemical Reactivity and Selectivity" 
by Jan Philipp Menzel, Benjamin B. Noble, James P. Blinco and Christopher Barner-Kowollik. 

Author: The code was created by Jan Philipp Menzel.
Contact by e-mail: j.p.menzel@gmail.com

Description for jpmenzelqled:

Purpose of the algorithm: Quantitative prediction of wavelength, photon number, time and concentration dependent conversion of photoreaction employing an LED.
Notes: The source code is designed to predict conversion of thioether-substituted o-methylbenzaldehyde A with N-ethylmaleiminde NEM (refer to the above-mentioned publication) 
using LED 2 (emission centered around 343 nm) in the respective 3D-printed photoreactor. 
Detailed Notes: The algorithm requests manual input, imports uv/vis data from a respective excel file, makes lists for spatial distribution of reactands and products, 
calculates time dependent development of overall conversion for wavelengths of the respective LED at requested amount of reactands as well as calculates light attenuation maps.
Different LEDs can be used for predictions and the respective data can be entered in the source code.

Description for jpmenzelqselectivity:

Purpose of the algorithm: Quantitative prediction of the wavelength-dependent selectivity of two competing photoreactions using monochromatic light.
The algorithm requests manual input, imports uv/vis data from a respective excel file, makes lists for spatial distribution of reactands and products, 
calculates time dependent development of overall conversion for varied monochromatic wavelengths at requested amount of reactands as well as calculates the required photon count.

Description for jpmenzelqledorthogonal:

Purpose of the algorithm: Quantitative prediction of wavelength, photon number, time and concentration dependent conversion of photoreaction employing LEDs.
This algorithm simulates the competing photoreaction between Dodecyl-thioether o-methylbenzaldehyde and a diaryltetrazole (N-Phenyl-p-OMe / C-Phenyl-p-methylester).
The algorithm requests manual input, imports uv/vis data from a respective excel file, makes lists for spatial distribution of reactands and products, 
calculates time dependent development of overall conversion for wavelengths of the respective LED at requested amount of reactands as well as calculates light attenuation maps.


