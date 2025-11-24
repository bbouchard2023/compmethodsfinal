% BMI Calculation, classification, and information storage/updating
% Must be called in one of two ways: 
% addVisitToExcel("PatientName.xlsx",height, weight, temperature, systolic bp, diastolic bp, Date w/Year-Month-Day Format");
% OR
% filename = ("PatientName.xlsx");
% height = height in m;
% weight = weight in kg;
% temp = temperature in °F
% sys_bp = blood pressure during peak systole in mmHg
% dia_bp = blood pressure during diastole in mmHg
% VisitDate = "Date in Year-Month-Day Format";
% addVisitToExcel(filename, height, weight, temperature, sys_bp, dia_bp, visitDate);
 
function addVisitToExcel(filename, height, weight, temp, sys_bp, dia_bp, visitDate)

    % ------------ BMI Calculation ------------
    BMI = weight / (height^2);

    if BMI < 18.5
            category = "Underweight";
        elseif BMI < 25
            category = "Optimal";
        elseif BMI < 30
            category = "Overweight";
        else
            category = "Obese";
    end
    
    % -------------- Body Temperature --------------
    if temp > 100.4
        tempcat = "Fever";
    elseif temp < 95
        tempcat = "Hypothermia";
    else 
        tempcat = "Normal Temperature";
    end
    
    % ------------ Blood Pressure Categorization ------------
    if sys_bp >= 160
            bp = "Stage 2 Hypertension";
    elseif sys_bp >= 140
        if dia_bp >= 100
            bp = "Stage 2 Hypertension";
        else 
            bp = "Stage 1 Hypertension";
        end
    elseif sys_bp >= 120
        if dia_bp >= 100 
            bp = "Stage 2 Hypertension";
        elseif dia_bp >= 90
            bp = "Stage 1 Hypertension";
        else 
            bp = "Pre Hypertension";
        end
    elseif sys_bp >= 90
        if dia_bp >= 100
            bp = "Stage 2 Hypertension";
        elseif dia_bp >= 90
            bp = "Stage 1 Hypertension";
        elseif dia_bp >= 80 
            bp = "Pre Hypertension";
        else
            bp = "Normal";
        end
    elseif sys_bp < 90
        if dia_bp >= 100
            bp = "Stage 2 Hypertension";
        elseif dia_bp >= 90
            bp = "Stage 1 Hypertension";
        elseif dia_bp >= 80
            bp = "Pre Hypertension";
        elseif dia_bp >= 60
            bp = "Normal";
        else
            bp = "Low";
        end
    end
    dbp = ((2 * dia_bp) + sys_bp) / 3;
    % ------------ Load or create Visits sheet ------------
    try
        opts = detectImportOptions(filename, 'Sheet', 'Visits');  % opens excel file
        visits = readtable(filename, opts);                       % finds sheet named "Visit"
        newRow = false;                                           % Indicates exisiting file is being updated
    catch
        % If the file or sheet doesn't exist, start a new one
        visits = table();                                         
        newRow = true;
    end

    % ------------ Append the new visit ------------
    newEntry = table( ...
        string(visitDate), height, weight, BMI, string(category), temp, string(tempcat), dbp, string(bp), ...
        'VariableNames', {'Date','Height_m','Weight_kg','BMI','BMI_Category','Temperature','Temp_Category','Blood Pressure','Blood Pressure Category'} ...
    );

    visits = [visits; newEntry];

    % ------------ Write back to Excel ------------
    writetable(visits, filename, 'Sheet', 'Visits', 'WriteMode', 'overwrite');


    % If profile sheet doesn’t exist, give the user a reminder
    if newRow
        fprintf('File created. Be sure to add patient name + DOB in the "Profile" sheet.\n');
    end

    fprintf('Visit recorded for file: %s\n', filename);
end