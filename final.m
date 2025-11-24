%==========================================================================
%% Comp Methods Final Project
%                                                             _,     ,_
%  Brendan Bouchard, Joe Manfre, Panagiotis Vasileidis      .'/  ,_   \'.
%  20251119                                                |  \__( >__/  |
%  Last Modified: 20251124                                 \             /
%                                                           '-..__ __..-'
%                                                           jgs  /_\
%==========================================================================

%% Front End
% view or input pt data


%% Biomarker Evaluation
% input pt biomarkers, evaluate good/bad

height = input("Height (m)? ");
weight = input("Weight (kg)? ");
temperature = input("Body Temperature (°F)? ");
sys_bp = input("Systolic Blood Pressure (mmHg)? ");
dia_bp = input("Diastolic Blood Pressure (mmHg)? ");
date = input("Today's Date (MMDDYYYY)? ");

%% Data Storage
% store pt biomarkers, generated data

addVisitToExcel(filename, height, weight, temperature, sys_bp, dia_bp, date)

%% Graphics
% display pt data, trends, etc from view pt data



%% Future Trend Approximation
% create approximated trend based on available data