# Machine-Learning-and-Data-Analytics-based-Operation-for-Smart-Grids

# Purpose/ Project Objectives

Smart grids show the potential to make an improvement in the power sector. By leveraging the knowledge of operation in a power system, which includes generation, distribution, consumption and two-way communication, machine learning and data analytics are utilized to improve the efficiency, reliability and reduce operational costs. 

In this project, the group seeks to leverage on machine learning and data analytic tools to gain insight and enhance performance of the smart distribution system with renewables. Based on a bottom-up data approach, it was determined that a possible way of improving the power grid notwithstanding existing constraints is an algorithm to increase cost effectiveness of power consumers during periods of high-power loss.

This is done through the calculation of each individual building’s power loss and reducing them during periods of high loss, with batteries charged by solar panels. In layman terms, during periods of high-power loss identified through the project’s neural network study, supply power would be limited and supplemented by batteries to meet the power demands of consumers, in turn reducing consumer costs and power loss during transmission.

The deliverables in the project includes the creation of a system network, correlation study between building consumption and irradiance data, neural network study and the established algorithm.



# Project Summary

Despite the lack of NTU’s network topology due to security constraints, assumptions were made, and the creation of a system network was successfully carried out. This set-in motion successful power flow within MATLAB, generating the necessary data for the calculation of power losses during transmissions within the network.

The project was also successful in implementing a solution based on batteries alongside each building’s solar panels to minimize power losses at any given time. It must be highlighted that the batteries used were not imposed any constraint. 

The analysis and visualization of data, including correlation between the 13 building’s consumption and irradiance, provided the insights necessary to the variables to be used in a machine learning model load forecasting. 

The goal of forecasting consumption for NTU buildings was partially successful. A neural network model was trained and is able to forecast real power consumption of 1 building based on the hour of the day, day of the week, month of the year and irradiance data, with an R2 score that shows substantial correlation between the building’s consumption and the forecasted data.  However, this was done for one building only due to time and computational constrains, and a complete study would require forecasting for all NTU buildings whose data is provided.

While the forecasts and power loss results are not integrated to deliver a concluding result within the project, it is general knowledge that periods of high building consumption would result in higher power losses during transmission. The model, albeit not having forecasted data due to time constraints, can be used as a reference upon completion for the scheduling of charging and discharging periods of the batteries implemented, reducing power losses.

The project’s initial goal of optimization and scheduling PV’s to match periods of high consumption were abolished for the reason that the optimal power flow failed to converge, possibly as a result of the utilization of case33BW in MATLAB, which is based on a different network topology as NTU. The plan for the increase in cost effectiveness for consumers through the reduction of power loss was adopted based on a bottom-up data approach. 


Done by: Nunes Di Pierri Enrique Alejandro, Lim Chee Quan Kelvin, Yang Ziqin, Tiah Jye Chen, Thniah Zhi Yang Bryan, Chua Ding Zhang
