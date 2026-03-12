%close all;
%clc;
%clear;

% RW DATA ANALYSIS MATLAB CODE
% Created by Luca Trussardi in May 2025

% The purpose of this code is to analyze test data obtained using the NEWARE cell tester BTS-4008-5V20A

% FIRST OF ALL, THE INPUT DATA NEEDS TO BE PLACED IN THE SAME DIRECTORY AS THE MATLAB FILE
% THE DATA SHOULD BE OBTAINED ACCORDING TO THE FOLLOWING EXPORT PROCEDURE:
%   FOR NEWARE (Voltages, Currents, ...) ----------------------------
% Once the test is completed, export the data from NEWARE BTSDA software (orange icon) using this settings
%   - Export report
%   - Export type: "Customize report"
%   - Export format: "EXCEL"
%   - Export templateConfiguration: "DefaultTemplate"
%   - Click "Export", a .xlsx file is produced
%   FOR DEWESOFT (Temperatures, ...) ----------------------------
% Once the test is completed, export the data from DewesoftX software (starting from the Data file) using this settings
%   - Export type: "EXCEL"
%   - Select the channels you want to export
%   - Click "Export"

% THEN, DATA SHOULD BE INPORTED IN THE CODE. THIS CAN BE DONE MANUALLY OR AUTOMATICALLY BY THE CODE.
% To automatically import data, select "Import" when the first window appears, and follow the instructions.
% Otherwise, to manually import data, follow this procedure:
%   FOR NEWARE (Voltages, Currents, ...) ----------------------------
%   - Import data using the Matlab "Import data" tool
%   - Select the Excel file containing the test data
%   - Select the "record" tab of the Excel sheet
%   - Set as output type "Table"
%   - Click on "Import selection"
%   - Close the "Import data" tool
%   FOR DEWESOFT (Temperatures, ...) ----------------------------
%   - Import data using the Matlab "Import data" tool
%   - Select the Excel file containing the test data
%   - Select the "Data1" tab of the Excel sheet
%   - Set as output type "Table"
%   - Click on "Import selection"
%   - Close the "Import data" tool


%% MAIN WINDOW

exit=false;

while not(exit) % Main window is continuously opened until one selects "Close"
    % Operation selection
    message='Cell Data Analysis Script';
    options=["1) Import NEWARE",...
             "2) Import DEWESOFT",...
             "3) Filter DEWESOFT",...
             "4) Plot NEWARE VS Time (single test, multiple quantities)",...
             "5) Plot DEWESOFT VS Time (single test, multiple quantities)",...
             "6) Plot NEWARE VS Time (multiple tests, single quantity)",...
             "7) Plot DEWESOFT VS Time (multiple tests, single quantity)",...
             "8) Plot NEWARE and DEWESOFT VS Time (single test, multiple quantities)",...
             "9) Export Data for SIMULINK",...
             "S) Save Data to .mat files",...
             "L) Load Data from .mat files",...
             "C) Close"];
    choice=menu(message,options); % Creates menu window that allows to select what to do

    switch choice
        case 1
            %% IMPORT NEWARE
            DataNEWARE=ImportFiles('record','DataNEWARE');

            % Fix Time, convert repeated 'hh:mm:ss' timestamps to seconds with 0.1s increments
            tableNames = fieldnames(DataNEWARE); % Get the available table names
                        
            for i = 1:length(tableNames) % Correct the time vector
                timeStrArray=DataNEWARE.(tableNames{i}).("Total Time");
                n = numel(timeStrArray);
                totalSeconds = zeros(n, 1);
                timeCount = containers.Map();
            
                for j = 1:n
                    tStr = timeStrArray{j};
            
                    if isKey(timeCount, tStr) % Count occurrences
                        timeCount(tStr) = timeCount(tStr) + 1;
                    else
                        timeCount(tStr) = 0;
                    end
            
                    timeParts = sscanf(tStr, '%d:%d:%d'); % Convert base time
                    baseSeconds = timeParts(1)*3600 + timeParts(2)*60 + timeParts(3);
            
                    totalSeconds(j) = baseSeconds + 0.1 * timeCount(tStr); % Count for tenth of seconds
                end

               DataNEWARE.(tableNames{i}).("Time(s)") = totalSeconds;
            end

        case 2
            %% IMPORT DEWESOFT
            DataDEWESOFT=ImportFiles('Data1','DataDEWESOFT','Data1-1');

            % Calculate Heat Flux from Voltage and automatically add it to data
            Sensitivity=17.39*10^(-6); % V*m^2/W
            TestArea=12.96*10^(-4); % m^2
            LateralSurfaceArea=2*pi*21*70*10^(-6); % m^2

            tableNames = fieldnames(DataDEWESOFT); % Get the available table names
            
            % For each of the tests, calculates the heat flux quantities and adds them to the data
            for i = 1:length(tableNames)          
                % Heat fluxes culculation with temperature corrected sensitivity value
                DataDEWESOFT.(tableNames{i}).("Heat Flux Density (W/m^2)")=DataDEWESOFT.(tableNames{i}).("Thermal flux FHF05")./(Sensitivity.*(1+0.002.*(DataDEWESOFT.(tableNames{i}).("Temperature FHF05")-20))); % W/m^2
                DataDEWESOFT.(tableNames{i}).("Heat Flux Test Area (W)")=DataDEWESOFT.(tableNames{i}).("Heat Flux Density (W/m^2)").*TestArea; % W on test area of the sensor
                DataDEWESOFT.(tableNames{i}).("Heat Flux Lateral Surface Area (W)")=DataDEWESOFT.(tableNames{i}).("Heat Flux Density (W/m^2)").*LateralSurfaceArea; % W on the lateral surface area

                % Rename 'Thermal flux FHF05' to 'Thermal Flux Sensor Voltage (V)'
                DataDEWESOFT.(tableNames{i}).Properties.VariableNames{'Thermal flux FHF05'} = 'Thermal Flux Sensor Voltage (V)';

                % Rename 'Temperature FHF05' to 'Temperature (C)'
                DataDEWESOFT.(tableNames{i}).Properties.VariableNames{'Temperature FHF05'} = 'Temperature (C)';
            end
          
        case 3
            %% FILTER DEWESOFT
            % Get the available table names
            tableNames = fieldnames(DataDEWESOFT);
            
            % Ask the user to select tables
            [selectedIdx, ok] = listdlg('ListString', tableNames, ...
                                        'SelectionMode', 'multiple', ...
                                        'PromptString', 'Select tables to filter:', ...
                                        'ListSize', [200 150]);
            
            % If the user cancels, exit
            if ~ok
                return;
            end
            
            % Extract the selected tables
            selectedTableNames = tableNames(selectedIdx);
            selectedTables = cell(size(selectedTableNames)); % Preallocate cell array
            for i = 1:length(selectedTableNames)
                selectedTables{i} = DataDEWESOFT.(selectedTableNames{i});
            end

            % Low pass filtering frequency requested as input
            prompt = {'Enter smoothdata filtering value:'};
            dlgtitle = 'Filtering';
            dims = [1 35]; % Size of the input field (rows x columns)
            defaultValue = {'1000'}; % Default value
            answer = inputdlg(prompt, dlgtitle, dims, defaultValue);
            smoothvalue = str2double(answer{1});

            % Filter data
            DataDEWESOFTFiltered = DataDEWESOFT;
            for i = 1:length(selectedTableNames)
                tableName = selectedTableNames{i};  % Get the table name
                T = DataDEWESOFT.(tableName);       % Get the corresponding table
                % Identify numeric columns, excluding 'Time'
                columnNames = T.Properties.VariableNames;
                numericCols = setdiff(columnNames, "Time");  % Exclude 'Time' column
                % Create a copy of the original table for filtered data
                TF = T;
                % Apply the filtering only to numeric columns
                for j = 1:length(numericCols)
                    colName = numericCols{j}; % Get column name
                    TF.(colName) = smoothdata(T.(colName), "movmean", smoothvalue); % Apply smoothing
                end
                % Store the filtered table
                DataDEWESOFTFiltered.(tableName) = TF;
            end

            % Create a figure for each filtered table
            for i = 1:length(selectedTableNames)
                tableName = selectedTableNames{i};  % Get the table name
                T = DataDEWESOFT.(tableName);       % Original data
                TF = DataDEWESOFTFiltered.(tableName); % Filtered data
                
                % Identify numeric columns excluding 'Time'
                columnNames = T.Properties.VariableNames;
                numericCols = setdiff(columnNames, "Time");  % Exclude 'Time'
                
                % Extract time vector
                TimeVector=(0:0.1:(height(T)-1)/10)';
                
                % Create figure for the current table
                figure('Name', tableName, 'NumberTitle', 'off');
                
                % Loop through each filtered variable and plot
                for j = 1:length(numericCols)
                    colName = numericCols{j}; % Get column name
                    
                    subplot(length(numericCols), 1, j); % Arrange subplots in a single column
                    hold on;
                    plot(TimeVector, T.(colName), 'b', 'LineWidth', 1.5, 'DisplayName', 'Original'); % Original data
                    plot(TimeVector, TF.(colName), 'r', 'LineWidth', 1.5, 'DisplayName', 'Filtered'); % Filtered data
                    hold off;
                    
                    xlabel('Time (s)');
                    ylabel(colName);
                    title(['Plot of ', colName, ' vs Time (s)']);
                    grid on;
                    legend('Location', 'best');
                end
            end
            
            % Accept filtering and save
            filterchoice=menu("Save filtered data?","Yes (Overwrites DataDEWSOFT with filtered data)","No (Does nothing)");
            
            switch filterchoice
                case 1 % Overvrites data with filtered one
                    DataDEWESOFT = DataDEWESOFTFiltered;
                case 2 % Nothing
            end

        case 4
            %% PLOT NEWARE VS TIME (single test, multiple quantities)
            % Get available table names
            tableNames = fieldnames(DataNEWARE);
            
            % Step 1: Let the user select a table
            [tableIndexN, okTable] = listdlg('ListString', tableNames, ...
                                            'SelectionMode', 'single', ...
                                            'PromptString', 'Select a table:', ...
                                            'ListSize', [200 150]);
            
            % Check if user made a selection
            if okTable == 0
                disp('No table selected. Exiting.');
                return;
            end
            
            % Get the selected table
            selectedTableNameN = tableNames{tableIndexN};
            selectedTableN = DataNEWARE.(selectedTableNameN);
            
            % Extract column names
            columnNamesN = selectedTableN.Properties.VariableNames;
            
            % Step 2: Let the user select columns to plot
            defaultSelectionN = [2,3,5,6];
            [selectedIndicesN, okColsN] = listdlg('ListString', columnNamesN, ...
                                                'SelectionMode', 'multiple', ...
                                                'InitialValue', defaultSelectionN, ... % Default selected values
                                                'PromptString', ['Select columns to plot from ', selectedTableNameN, ':'], ...
                                                'ListSize', [400 150]);
            
            % Check if user made a selection
            if okColsN == 0
                disp('No columns selected. Exiting.');
                return;
            end
            
            % Get the selected column names
            selectedColsN = columnNamesN(selectedIndicesN);
            
            % Create the time vector
            TimeVector=selectedTableN.('Time(s)');
            
            plotchoice=menu("How do you want to plot?","Single figure with subplots","Separate figures","Single plot with multiple axes");
            
            switch plotchoice
                case 1
                % Create a figure with subplots
                figure('Name', selectedTableNameN, 'NumberTitle', 'off');
                numPlots = length(selectedColsN);
                for i = 1:numPlots
                    subplot(numPlots, 1, i); % Arrange subplots in a single column
                    plot(TimeVector, selectedTableN{:, selectedColsN{i}}, 'LineWidth', 1.5);
                    xlabel('Time(s)');
                    ylabel(selectedColsN{i});
                    title(['Plot of ', selectedColsN{i}, ' vs Time(s)']);
                    grid on;
                end
                case 2
                % Create separate figures
                for i = 1:length(selectedColsN)
                    figure;
                    plot(TimeVector, selectedTableN{:, selectedColsN{i}}, 'LineWidth', 1.5);
                    xlabel('Time(s)');
                    ylabel(selectedColsN{i});
                    title(['Plot of ', selectedColsN{i}, ' vs Time(s)']);
                    grid on;
                end
                case 3
                    % Create a figure with multiple axes
                    figure;
                    ax1 = axes; % Create the first axes object
                    colors = lines(length(selectedColsN)); % Generate distinct colors

                    % Create an array to hold all axes
                    allAxes = ax1;

                    % Loop through selected columns and create multiple y-axes
                    hold on;
                    for i = 1:length(selectedColsN)
                        if i == 1
                            % First plot on main axes (ax1)
                            plot(ax1, TimeVector, selectedTableN{:, selectedColsN{i}}, 'Color', colors(i, :), 'LineWidth', 1.5);
                            ax1.YColor = colors(i, :);
                            ylabel(ax1, selectedColsN{i});
                        else
                            % Create new axes on top of existing one
                            axNew = axes('Position', ax1.Position, 'Color', 'none', 'YAxisLocation', 'right');
                            hold on;
                            plot(axNew, TimeVector, selectedTableN{:, selectedColsN{i}}, 'Color', colors(i, :), 'LineWidth', 1.5);
                            axNew.YColor = colors(i, :);
                            ylabel(axNew, selectedColsN{i});
                            set(axNew, 'XColor', 'none'); % Hide x-axis ticks on additional axes
                            allAxes = [allAxes, axNew]; % Add the new axis to the allAxes array
                        end
                    end
                    hold off;

                    % Formatting
                    xlabel(ax1, 'Time(s)');
                    title('Multi Y-axes Plot');
                    grid on;

                    % Link the axes to synchronize zooming
                    linkaxes(allAxes, 'x'); % Link all axes for x-axis zooming

                    % Optional: add legend
                    %legend(selectedColsN, 'Location', 'best');
            end

        case 5
            %% PLOT DEWESOFT VS TIME (single test, multiple quantities)
            % Get available table names
            tableNames = fieldnames(DataDEWESOFT);
            
            % Step 1: Let the user select a table
            [tableIndexD, okTable] = listdlg('ListString', tableNames, ...
                                            'SelectionMode', 'single', ...
                                            'PromptString', 'Select a table:', ...
                                            'ListSize', [200 150]);
            
            % Check if user made a selection
            if okTable == 0
                disp('No table selected. Exiting.');
                return;
            end
            
            % Get the selected table
            selectedTableNameD = tableNames{tableIndexD};
            selectedTableD = DataDEWESOFT.(selectedTableNameD);
            
            % Extract column names
            columnNamesD = selectedTableD.Properties.VariableNames;
            
            % Step 2: Let the user select columns to plot
            defaultSelectionD = [2,3,4,5,6];
            [selectedIndicesD, okColsD] = listdlg('ListString', columnNamesD, ...
                                                'SelectionMode', 'multiple', ...
                                                'InitialValue', defaultSelectionD, ... % Default selected values
                                                'PromptString', ['Select columns to plot from ', selectedTableNameD, ':'], ...
                                                'ListSize', [400 150]);
            
            % Check if user made a selection
            if okColsD == 0
                disp('No columns selected. Exiting.');
                return;
            end
            
            % Get the selected column names
            selectedColsD = columnNamesD(selectedIndicesD);
            
            % Create the time vector
            TimeVector=(0:0.1:(height(selectedTableD)-1)/10)';
            
            plotchoice=menu("How do you want to plot?","Single figure with subplots","Separate figures","Single plot with multiple axes");
            
            switch plotchoice
                case 1
                % Create a figure with subplots
                figure('Name', selectedTableNameD, 'NumberTitle', 'off');
                numPlots = length(selectedColsD);
                for i = 1:numPlots
                    subplot(numPlots, 1, i); % Arrange subplots in a single column
                    plot(TimeVector, selectedTableD{:, selectedColsD{i}}, 'LineWidth', 1.5);
                    xlabel('Time(s)');
                    ylabel(selectedColsD{i});
                    title(['Plot of ', selectedColsD{i}, ' vs Time(s)']);
                    grid on;
                end
                case 2
                % Create separate figures
                for i = 1:length(selectedColsD)
                    figure;
                    plot(TimeVector, selectedTableD{:, selectedColsD{i}}, 'LineWidth', 1.5);
                    xlabel('Time(s)');
                    ylabel(selectedColsD{i});
                    title(['Plot of ', selectedColsD{i}, ' vs Time(s)']);
                    grid on;
                end
                case 3
                    % Create a figure with multiple axes
                    figure;
                    ax1 = axes; % Create the first axes object
                    colors = lines(length(selectedColsD)); % Generate distinct colors

                    % Create an array to hold all axes
                    allAxes = ax1;

                    % Loop through selected columns and create multiple y-axes
                    hold on;
                    for i = 1:length(selectedColsD)
                        if i == 1
                            % First plot on main axes (ax1)
                            plot(ax1, TimeVector, selectedTableD{:, selectedColsD{i}}, 'Color', colors(i, :), 'LineWidth', 1.5);
                            ax1.YColor = colors(i, :);
                            ylabel(ax1, selectedColsD{i});
                        else
                            % Create new axes on top of existing one
                            axNew = axes('Position', ax1.Position, 'Color', 'none', 'YAxisLocation', 'right');
                            hold on;
                            plot(axNew, TimeVector, selectedTableD{:, selectedColsD{i}}, 'Color', colors(i, :), 'LineWidth', 1.5);
                            axNew.YColor = colors(i, :);
                            ylabel(axNew, selectedColsD{i});
                            set(axNew, 'XColor', 'none'); % Hide x-axis ticks on additional axes
                            allAxes = [allAxes, axNew]; % Add the new axis to the allAxes array
                        end
                    end
                    hold off;

                    % Formatting
                    xlabel(ax1, 'Time(s)');
                    title('Multi Y-axes Plot');
                    grid on;

                    % Link the axes to synchronize zooming
                    linkaxes(allAxes, 'x'); % Link all axes for x-axis zooming

                    % Optional: add legend
                    %legend(selectedColsN, 'Location', 'best');
            end

        case 6
            %% PLOT NEWARE SPECIFIC PROPERTY VS TIME (multiple tests, single quantity)
            % Get the available table names
            tableNames = fieldnames(DataNEWARE);
            
            % Ask the user to select tables
            [selectedIdx, ok] = listdlg('ListString', tableNames, ...
                                        'SelectionMode', 'multiple', ...
                                        'PromptString', 'Select tables to plot:', ...
                                        'ListSize', [200 150]);
            
            % If the user cancels, exit
            if ~ok
                return;
            end
            
            % Extract the selected tables
            selectedTables = tableNames(selectedIdx);
            
            % Get the common variable names (assuming all tables have the same columns)
            commonVars = DataNEWARE.(selectedTables{1}).Properties.VariableNames;
            
            % Ask the user to select a property to plot
            [propIdx, propOk] = listdlg('ListString', commonVars, ...
                                        'SelectionMode', 'single', ...
                                        'PromptString', 'Select the property to plot:');
            
            % If the user cancels, exit
            if ~propOk
                return;
            end
            
            selectedProperty = commonVars{propIdx};
            
            % Plot the selected property against time for all selected tables
            figure;
            hold on;
            for i = 1:length(selectedTables)
                T = DataNEWARE.(selectedTables{i});
                currentheight = height(T);
                
                % If the table is shorter than maxheight, extend it
                if currentheight < maxheight
                    numNewRows = maxheight - currentheight;
                    
                    % Create an empty table with the same variable names and matching data types
                    varTypes = varfun(@class, T, 'OutputFormat', 'cell'); % Get column types
                    zeroData = cell(1, width(T)); % Initialize empty cell array
            
                    for col = 1:width(T)
                        if strcmp(varTypes{col}, 'double') % Numeric columns
                            zeroData{col} = zeros(numNewRows, 1);
                        elseif strcmp(varTypes{col}, 'cell') % Cell array columns
                            zeroData{col} = repmat({''}, numNewRows, 1);
                        elseif strcmp(varTypes{col}, 'string') % String columns
                            zeroData{col} = strings(numNewRows, 1);
                        else
                            warning('Unsupported data type in column "%s".', T.Properties.VariableNames{col});
                        end
                    end
                    
                    % Convert to table
                    zeroRows = table(zeroData{:}, 'VariableNames', T.Properties.VariableNames);
                    
                    % Append zeroRows to the original table
                    T = [T; zeroRows];
                end
                
                % Plot
                plot(T.('Time(s)'), T.(selectedProperty), 'DisplayName', selectedTables{i}, 'LineWidth', 1.5);
            end

            % Formatting
            xlabel('Time');
            ylabel(selectedProperty);
            title(['Plot of ', selectedProperty]);
            leg=legend;
            set(leg,'Interpreter', 'none');
            grid on;
            hold off;

        case 7
            %% PLOT DEWESOFT SPECIFIC PROPERTY VS TIME (multiple tests, single quantity)
            % Get the available table names
            tableNames = fieldnames(DataDEWESOFT);
            
            % Ask the user to select tables
            [selectedIdx, ok] = listdlg('ListString', tableNames, ...
                                        'SelectionMode', 'multiple', ...
                                        'PromptString', 'Select tables to plot:', ...
                                        'ListSize', [200 150]);
            
            % If the user cancels, exit
            if ~ok
                return;
            end
            
            % Extract the selected tables
            selectedTables = tableNames(selectedIdx);
            
            % Get the common variable names (assuming all tables have the same columns)
            commonVars = DataDEWESOFT.(selectedTables{1}).Properties.VariableNames;
            
            % Ask the user to select a property to plot
            [propIdx, propOk] = listdlg('ListString', commonVars, ...
                                        'SelectionMode', 'single', ...
                                        'PromptString', 'Select the property to plot:');
            
            % If the user cancels, exit
            if ~propOk
                return;
            end
            
            selectedProperty = commonVars{propIdx};
            
            % Create time vector
            maxheight = 0;
            for i = 1:length(selectedTables)
                T = DataDEWESOFT.(selectedTables{i});
                if height(T) > maxheight
                    maxheight = height(T);
                end   
            end
            TimeVector = (0:0.1:(maxheight-1)/10)';
            
            % Plot the selected property against time for all selected tables
            figure;
            hold on;
            for i = 1:length(selectedTables)
                T = DataDEWESOFT.(selectedTables{i});
                currentheight = height(T);
                
                % If the table is shorter than maxheight, extend it
                if currentheight < maxheight
                    numNewRows = maxheight - currentheight;
                    
                    % Create an empty table with the same variable names and matching data types
                    varTypes = varfun(@class, T, 'OutputFormat', 'cell'); % Get column types
                    zeroData = cell(1, width(T)); % Initialize empty cell array
            
                    for col = 1:width(T)
                        if strcmp(varTypes{col}, 'double') % Numeric columns
                            zeroData{col} = NaN(numNewRows, 1);
                        elseif strcmp(varTypes{col}, 'cell') % Cell array columns
                            zeroData{col} = repmat({''}, numNewRows, 1);
                        elseif strcmp(varTypes{col}, 'string') % String columns
                            zeroData{col} = strings(numNewRows, 1);
                        else
                            warning('Unsupported data type in column "%s".', T.Properties.VariableNames{col});
                        end
                    end
                    
                    % Convert to table
                    zeroRows = table(zeroData{:}, 'VariableNames', T.Properties.VariableNames);
                    
                    % Append zeroRows to the original table
                    T = [T; zeroRows];
                end
                
                % Plot
                plot(TimeVector, T.(selectedProperty), 'DisplayName', selectedTables{i}, 'LineWidth', 1.5);
            end

            % Formatting
            xlabel('Time');
            ylabel(selectedProperty);
            title(['Plot of ', selectedProperty]);
            leg=legend;
            set(leg,'Interpreter', 'none');
            grid on;
            hold off;

        case 8
            %% PLOT NEWARE AND DEWESOFT VS TIME (single test, multiple quantities)
            % Get available table names
            tableNamesNEWARE = fieldnames(DataNEWARE);
            tableNamesDEWESOFT = fieldnames(DataDEWESOFT);
            
            % Step 1a: Let the user select a NEWARE table
            [tableIndexN, okTableN] = listdlg('ListString', tableNamesNEWARE, ...
                                            'SelectionMode', 'single', ...
                                            'PromptString', 'Select a NEWARE table:', ...
                                            'ListSize', [200 150]);
            
            % Check if user made a selection
            if okTableN == 0
                disp('No table selected. Exiting.');
                return;
            end
            
            % Step 1b: Let the user select a DEWESOFT table
            [tableIndexD, okTableD] = listdlg('ListString', tableNamesDEWESOFT, ...
                                            'SelectionMode', 'single', ...
                                            'PromptString', 'Select a DEWESOFT table:', ...
                                            'ListSize', [200 150]);
            
            % Check if user made a selection
            if okTableD == 0
                disp('No table selected. Exiting.');
                return;
            end

            % Get the selected tables
            selectedTableNameN = tableNamesNEWARE{tableIndexN};
            selectedTableN = DataNEWARE.(selectedTableNameN);
            selectedTableNameD = tableNamesDEWESOFT{tableIndexD};
            selectedTableD = DataDEWESOFT.(selectedTableNameD);
            
            % Extract column names
            columnNamesN = selectedTableN.Properties.VariableNames;
            columnNamesD = selectedTableD.Properties.VariableNames;
            
            % Step 2a: Let the user select NEWARE columns to plot
            defaultSelectionN = [];
            [selectedIndicesN, okColsN] = listdlg('ListString', columnNamesN, ...
                                                'SelectionMode', 'multiple', ...
                                                'InitialValue', defaultSelectionN, ... % Default selected values
                                                'PromptString', ['Select NEWARE columns to plot from ', selectedTableNameN, ':'], ...
                                                'ListSize', [400 150]);
            
            % Check if user made a selection
            if okColsN == 0
                disp('No columns selected. Exiting.');
                return;
            end

            % Step 2b: Let the user select DEWESOFT columns to plot
            defaultSelectionD = [];
            [selectedIndicesD, okColsD] = listdlg('ListString', columnNamesD, ...
                                                'SelectionMode', 'multiple', ...
                                                'InitialValue', defaultSelectionD, ... % Default selected values
                                                'PromptString', ['Select DEWESOFT columns to plot from ', selectedTableNameD, ':'], ...
                                                'ListSize', [400 150]);
            
            % Check if user made a selection
            if okColsD == 0
                disp('No columns selected. Exiting.');
                return;
            end
            
            % Get the selected column names
            selectedColsN = columnNamesN(selectedIndicesN);
            selectedColsD = columnNamesD(selectedIndicesD);
            
            % Create the time vector
            TimeVectorN=selectedTableN.('Time(s)');
            TimeVectorD=(0:0.1:(height(selectedTableD)-1)/10)';
            if TimeVectorN(end)>TimeVectorD(end)
                TimeVector=(0:0.1:TimeVectorN(end))';
            else
                TimeVector=(0:0.1:TimeVectorD(end))';
            end
            
            plotchoice=menu("How do you want to plot?","Single figure with subplots","Separate figures");
            
            switch plotchoice
                case 1
                % Create a figure with subplots
                figure('Name', [selectedTableNameN, ' and ', selectedTableNameD], 'NumberTitle', 'off');
                numPlots = length(selectedColsN)+length(selectedColsD);
                for i = 1:length(selectedColsN)
                    subplot(numPlots, 1, i); % Arrange subplots in a single column
                    plot(TimeVectorN, selectedTableN{:, selectedColsN{i}}, 'LineWidth', 1.5);
                    xlabel('Time(s)');
                    ylabel(selectedColsN{i});
                    title(['Plot of ', selectedColsN{i}, ' vs Time(s)']);
                    grid on;
                end
                for i = 1:length(selectedColsD)
                    subplot(numPlots, 1, i+length(selectedColsN)); % Arrange subplots in a single column
                    plot(TimeVectorD, selectedTableD{:, selectedColsD{i}}, 'LineWidth', 1.5);
                    xlabel('Time(s)');
                    ylabel(selectedColsD{i});
                    title(['Plot of ', selectedColsD{i}, ' vs Time(s)']);
                    grid on;
                end
                case 2
                % Create separate figures
                for i = 1:length(selectedColsN)
                    figure;
                    plot(TimeVectorN, selectedTableN{:, selectedColsN{i}}, 'LineWidth', 1.5);
                    xlabel('Time(s)');
                    ylabel(selectedColsN{i});
                    title(['Plot of ', selectedColsN{i}, ' vs Time(s)']);
                    grid on;
                end
                for i = 1:length(selectedColsD)
                    figure;
                    plot(TimeVectorD, selectedTableD{:, selectedColsD{i}}, 'LineWidth', 1.5);
                    xlabel('Time(s)');
                    ylabel(selectedColsD{i});
                    title(['Plot of ', selectedColsD{i}, ' vs Time(s)']);
                    grid on;
                end
            end

        case 9
            %% EXPORT DATA FOR SIMULINK
            % Get available table names
            tableNamesNEWARE = fieldnames(DataNEWARE);
            tableNamesDEWESOFT = fieldnames(DataDEWESOFT);
            
            % Step 1a: Let the user select a NEWARE table
            [tableIndexN, okTableN] = listdlg('ListString', tableNamesNEWARE, ...
                                            'SelectionMode', 'single', ...
                                            'PromptString', 'Select a NEWARE table:', ...
                                            'ListSize', [200 150]);
            
            % Check if user made a selection
            if okTableN == 0
                disp('No table selected. Exiting.');
                return;
            end

            % Step 1b: Let the user select a DEWESOFT table
            [tableIndexD, okTableD] = listdlg('ListString', tableNamesDEWESOFT, ...
                                            'SelectionMode', 'single', ...
                                            'PromptString', 'Select a DEWESOFT table:', ...
                                            'ListSize', [200 150]);
            
            % Check if user made a selection
            if okTableD == 0
                disp('No table selected. Exiting.');
                return;
            end

            % Get the selected tables
            selectedTableNameN = tableNamesNEWARE{tableIndexN};
            selectedTableN = DataNEWARE.(selectedTableNameN);
            selectedTableNameD = tableNamesDEWESOFT{tableIndexD};
            selectedTableD = DataDEWESOFT.(selectedTableNameD);

            % Export Current
            Current_Profile = timeseries(selectedTableN.('Current(A)'), selectedTableN.('Time(s)'));
            save('Current_Profile.mat', 'Current_Profile');

            % Export Voltage
            Voltage_Profile = timeseries(selectedTableN.('Voltage(V)'), selectedTableN.('Time(s)'));
            save('Voltage_Profile.mat', 'Voltage_Profile');

            % Export Temperature
            Temperature_Profile = timeseries(selectedTableD.('Temperature (C)'), selectedTableD.('Time'));
            save('Temperature_Profile.mat', 'Temperature_Profile');

            % Export Heat Flux Density
            Heat_Flux_Density_Profile = timeseries(selectedTableD.('Heat Flux Density (W/m^2)'), selectedTableD.('Time'));
            save('Heat_Flux_Density_Profile.mat', 'Heat_Flux_Density_Profile');

            disp('Success: Data exported to .mat files');

           case 10
            %% SAVE DATA TO .MAT FILES
            % Check if 'DataNEWARE' exists in the workspace and save it
            if exist('DataNEWARE', 'var')
                save('DataNEWARE.mat', 'DataNEWARE');
                fprintf('DataNEWARE saved to DataNEWARE.mat\n');
            end
            
            % Check if 'DataDEWESOFT' exists in the workspace and save it
            if exist('DataDEWESOFT', 'var')
                save('DataDEWESOFT.mat', 'DataDEWESOFT');
                fprintf('DataDEWESOFT saved to DataDEWESOFT.mat\n');
            end
            
            % Display a message if neither variable exists
            if ~exist('DataNEWARE', 'var') && ~exist('DataDEWESOFT', 'var')
                fprintf('No variables found to save.\n');
            end

        case 11
            %% LOAD DATA FROM .MAT FILES
            % Check if 'DataNEWARE.mat' exists and load it
            if exist('DataNEWARE.mat', 'file')
                load('DataNEWARE.mat');
                fprintf('DataNEWARE loaded from DataNEWARE.mat\n');
            end
            
            % Check if 'DataDEWESOFT.mat' exists and load it
            if exist('DataDEWESOFT.mat', 'file')
                load('DataDEWESOFT.mat');
                fprintf('DataDEWESOFT loaded from DataDEWESOFT.mat\n');
            end
            
            % Display a message if neither file exists
            if ~exist('DataNEWARE.mat', 'file') && ~exist('DataDEWESOFT.mat', 'file')
                fprintf('No .mat files found to load.\n');
            end

        case 12
            %% CLOSE
            exit=true;
        otherwise
            exit=true;
    end

end
