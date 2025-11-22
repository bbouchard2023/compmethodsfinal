% BMI Calculation, classification, and information storage/updating
% Must be called in one of two ways: 
% addVisitToExcel("PatientName.xlsx",height, weight,"Date w/Year-Month-Day Format");
% OR
% filename = ("PatientName.xlsx");
% height = height in m;
% weight = weight in kg;
% VisitDate = "Date in Year-Month-Day Format";
% addVisitToExcel(filename, height, weight, visitDate);
 
function addVisitToExcel(filename, height, weight, visitDate)

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
        string(visitDate), height, weight, BMI, string(category), ...
        'VariableNames', {'Date','Height_m','Weight_kg','BMI','BMI_Category'} ...
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