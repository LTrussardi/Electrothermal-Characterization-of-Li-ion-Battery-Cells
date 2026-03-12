%close all;
%clc;
%clear;

% HPPC DATA ANALYSIS MATLAB CODE
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
             "9) Create HPPC Data (MANUAL CODE)",...
             "9E) Export MANUAL Results to Excel file",...
             "10) Create HPPC Data (MATLAB AUTO CODE)",...
             "10E) Export AUTO Results to Excel file",...
             "11) Compare HPPC Results",...
             "12) Estimate Heat Capacity", ...
             "13) Export Data for SIMULINK",...
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

            % Calculate SOC
            for i = 1:length(tableNames)
                T=DataNEWARE.(tableNames{i});
                % Finding the real Capacity curve by integration (the one provided is reset at each phase)
                Capacity=cumtrapz(DataNEWARE.(tableNames{i}).("Time(s)"),DataNEWARE.(tableNames{i}).("Current(A)"))/3600;
                T.('Real Capacity by Integration(Ah)')=Capacity;
                % Finding the Capacity corresponding to 100% SOC (Start of pilot discharge)
                cccv_idx = strcmp(T.('Step Type'), 'CCCV Chg'); % Identify where 'Step Type' is 'CCCV Chg'              
                transitions = diff([0; cccv_idx; 0]); % Find start and end indices of all continuous blocks
                starts = find(transitions == 1);
                ends   = find(transitions == -1) - 1;
                if length(starts) < 2 % Check if at least 2 blocks exist
                    disp('Less than two CCCV Chg blocks found.');
                    Capacity100 = NaN;
                else
                    second_block_end = ends(2); % Get the end index of the second block
                    Capacity100 = T.('Real Capacity by Integration(Ah)')(second_block_end); % Retrieve the corresponding 'Real Capacity by Integration (Ah)'
                end
                % Finding the Capacity corresponding to 0% SOC (End of pilot discharge)
                rest_idx = strcmp(T.('Step Type'), 'Rest'); % Identify where 'Step Type' is 'Rest'
                transitions = diff([0; rest_idx; 0]); % Find start and end indices of all continuous blocks
                starts = find(transitions == 1);
                ends   = find(transitions == -1) - 1;
                if isempty(starts) % Check if any blocks exist
                    disp('No Rest blocks found.');
                    Capacity0 = NaN;
                else
                    last_block_end = ends(end); % Get the end index of the last block
                    Capacity0 = T.('Real Capacity by Integration(Ah)')(last_block_end); % Retrieve the corresponding 'Real Capacity by Integration (Ah)'
                end                
                % Finding maximum capacity during the pilot discharge phase (used as reference)
                maxCapacity=Capacity100-Capacity0;
                % Finding DOD and SOC
                T.('SOC')=(T.('Real Capacity by Integration(Ah)')-Capacity0)./maxCapacity.*100;
                DataNEWARE.(tableNames{i})=T;
                DataNEWARE.(tableNames{i}).('Pilot discharge capacity (Ah)')(:)=maxCapacity; % Save value
            end

            % Delete useless coloumns and rename some of them for more clarity
            for i = 1:length(tableNames)
                T=DataNEWARE.(tableNames{i});
                T(:, {'DataPoint', 'Time','Total Time','Capacity(Ah)','Energy(Wh)','Date','Power(W)'}) = [];
                T.Properties.VariableNames{strcmp(T.Properties.VariableNames, 'Real Capacity by Integration(Ah)')} = 'Capacity(Ah)';
                DataNEWARE.(tableNames{i})=T;
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
            %% CREATE HPPC DATA (MANUAL CODE)
            % Unite NEWARE and DEWESOFT data
            % Initialize the combined struct (merge of NEWARE and DEWESOFT Data)
            DataCOMBINED = struct();
            DataHPPC = struct();

            % Get the field names of both structs
            fieldsNEWARE = fieldnames(DataNEWARE);
            fieldsDEWESOFT = fieldnames(DataDEWESOFT);

            % Loop over the fields of DataNEWARE to find corresponding fields in DataDEWESOFT
            for i = 1:length(fieldsNEWARE)
                % Check if the current field name ends with 'NEWARE'
                if endsWith(fieldsNEWARE{i}, '_NEWARE_xlsx')
                    % Extract the base name by removing 'NEWARE'
                    baseName = extractBefore(fieldsNEWARE{i}, '_NEWARE_xlsx');
                    % Check if the same base name exists in DataDEWESOFT with 'DEWESOFT'
                    correspondingField = strcat(baseName, '_DEWESOFT_xlsx');
                    if isfield(DataDEWESOFT, correspondingField)
                        % Get the tables from both structs
                        tableNEWARE = DataNEWARE.(fieldsNEWARE{i});
                        tableDEWESOFT = DataDEWESOFT.(correspondingField);

                        % Rename 'Time' to 'Time(s)' for DEWESOFT
                        tableDEWESOFT.Properties.VariableNames{strcmp(tableDEWESOFT.Properties.VariableNames, 'Time')} = 'Time(s)';
                        
                        % Merge the tables using an outer join on 'Time(s)'
                        combinedTable = outerjoin(tableNEWARE, tableDEWESOFT, 'Keys', 'Time(s)', 'MergeKeys', true);

                        % Fill Nan and empty strings with data
                        filledTable = combinedTable;
                        varNames = filledTable.Properties.VariableNames;
                    
                        for k = 1:numel(varNames)
                            colData = filledTable.(varNames{k});
                    
                            % Handle numeric columns
                            if isnumeric(colData)
                                % Forward fill
                                lastVal = NaN;
                                for j = 1:length(colData)
                                    if ~isnan(colData(j))
                                        lastVal = colData(j);
                                    elseif ~isnan(lastVal)
                                        colData(j) = lastVal;
                                    end
                                end
                                % Fill top NaNs with first valid value
                                if isnan(colData(1))
                                    firstValid = find(~isnan(colData), 1, 'first');
                                    if ~isempty(firstValid)
                                        colData(1:firstValid-1) = colData(firstValid);
                                    end
                                end
                                filledTable.(varNames{k}) = colData;
                    
                            % Handle string/cell string columns
                            elseif iscellstr(colData) || isstring(colData)
                                % Convert to cellstr if it's string for uniform handling
                                if isstring(colData)
                                    colData = cellstr(colData);
                                    convertBackToString = true;
                                else
                                    convertBackToString = false;
                                end
                    
                                lastStr = '';
                                for j = 1:length(colData)
                                    if ~strcmp(colData{j}, '')
                                        lastStr = colData{j};
                                    elseif ~strcmp(lastStr, '')
                                        colData{j} = lastStr;
                                    end
                                end
                    
                                % Fill top empty strings with first non-empty
                                if strcmp(colData{1}, '')
                                    firstValid = find(~cellfun(@isempty, colData), 1, 'first');
                                    if ~isempty(firstValid)
                                        colData(1:firstValid-1) = repmat(colData(firstValid), firstValid-1, 1);
                                    end
                                end
                    
                                % Convert back if necessary
                                if convertBackToString
                                    colData = string(colData);
                                end
                    
                                filledTable.(varNames{k}) = colData;
                            end
                        end

                        % Keep only HPPC test (delete inititial charge/discharge) by starting to consider data from the 4th 'Rest' phase
                        is_rest = strcmp(filledTable.('Step Type'), 'Rest'); % Logical index of 'Rest' rows
                        transitions = diff([0; is_rest; 0]);  % Find the start indices of all 'Rest' blocks
                        starts = find(transitions == 1);      % start of each block
                        ends   = find(transitions == -1) - 1; % end of each block (not used here)
                        if length(starts) < 4 % Check if there are at least 4 'Rest' blocks
                            disp('Less than 4 ''Rest'' blocks found.');
                            trimmedTable = table();  % return empty table
                        else
                            start_idx = starts(4); % Get the start index of the 4th block
                            trimmedTable = filledTable(start_idx:end, :); % Keep the table from that point onward
                            trimmedTable.('Time(s)') = trimmedTable.('Time(s)')-trimmedTable.('Time(s)')(1); % Adjust time
                        end

                        % Keep only HPPC test (delete final discharge) by finishing to consider up to the second to last 'Rest' phase
                        is_rest = strcmp(trimmedTable.('Step Type'), 'Rest'); % Logical index of 'Rest' rows
                        transitions = diff([0; is_rest; 0]); % Identify block starts and ends
                        starts = find(transitions == 1);
                        ends   = find(transitions == -1) - 1;
                        if length(starts) < 2 % Ensure there are at least 2 blocks
                            disp('Less than two ''Rest'' blocks found.');
                            trimmedTable2 = trimmedTable;  % return original table
                        else
                            cut_idx = ends(end-1);  % Get the end index of the second-to-last block  
                            trimmedTable2 = trimmedTable(1:cut_idx, :); % Keep only rows up to that point
                        end

                        % Save the combined table in the new struct
                        DataCOMBINED.(baseName) = trimmedTable2;

                        % Save the combined table in the new struct (Only useful data)
                        DataHPPC.(baseName) = DataCOMBINED.(baseName)(:, {'Step Type','Time(s)', 'Current(A)', 'Voltage(V)', 'Temperature (C)','SOC'});
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Time(s)')} = 'Time (s)';
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Current(A)')} = 'Current (A)';
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Voltage(V)')} = 'Voltage (V)';
                    end
                end
            end

            % Retrieving HPPC parameters
            fields=fieldnames(DataHPPC);
            OCV=struct();
            R0=struct();
            R1=struct();
            C1=struct();
            R2=struct();
            C2=struct();
            Discharge=struct();
            DischargeRelaxation=struct();
            Charge=struct();
            ChargeRelaxation=struct();
            a=struct();
            b=struct();
            c=struct();
            tau1=struct();
            tau2=struct();
            V1=struct();
            V2=struct();

            for i = 1:length(fields) % Iterating all the tests
                Data=DataHPPC.(fields{i});
                
                % Identifying rests
                is_rest = strcmp(Data.('Step Type'), 'Rest');
                transitions = diff([0; is_rest; 0]);
                starts = find(transitions == 1);
                ends = find(transitions == -1) - 1;
                numSteps = floor(length(starts)/3);

                DischargeArray=cell(numSteps,1);
                DischargeRelaxationArray=cell(numSteps,1);
                ChargeArray=cell(numSteps,1);
                ChargeRelaxationArray=cell(numSteps,1);

                % Find OCV at 10% SOC increments (voltage at the end of the 1h rests, that are the first, fourth, seventh, ...)
                OCVdata=table('Size',[numSteps 2],'VariableNames',["OCV","SOC"],'VariableTypes',{'double','double'});
                for j=1:numSteps % Iterating the SOC increments
                    OCVdata.OCV(j)=Data.('Voltage (V)')(ends(1+(j-1)*3));
                    OCVdata.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                end
                OCV.(fields{i})=OCVdata;

                % Finding R0 at 10% SOC increments (instantaneous voltage drop)
                R0data=table('Size',[numSteps 7],'VariableNames',["R0Discharge","R0DischargeRelaxation","R0Charge","R0ChargeRelaxation","R0DischargeAVG","R0ChargeAVG","SOC"],'VariableTypes',{'double','double','double','double','double','double','double'});
                for j=1:numSteps % Iterating the SOC increments
                    R0data.R0Discharge(j)=(OCVdata.OCV(j)-Data.('Voltage (V)')(ends(1+(j-1)*3)+1))/abs(Data.('Current (A)')(ends(1+(j-1)*3)+1)); % R0 D=(VA-VB)/Id, where VA=VOC (voltage at the end of the first rest) and VB is the voltage sampled immediately after
                    R0data.R0DischargeRelaxation(j)=(Data.('Voltage (V)')(starts(2+(j-1)*3))-Data.('Voltage (V)')(starts(2+(j-1)*3)-1))/abs(Data.('Current (A)')(starts(2+(j-1)*3)-1)); % R0 DR=(VD-VC)/Id, where VD is the voltage at the start of the second rest and VC the voltage sampled immediately before
                    R0data.R0Charge(j)=(Data.('Voltage (V)')(ends(2+(j-1)*3)+1)-Data.('Voltage (V)')(ends(2+(j-1)*3)))/abs(Data.('Current (A)')(ends(2+(j-1)*3)+1)); % R0 C=(VF-VE)/Ic, where VF is the voltage after the end of the second rest and VE the voltage sampled immediately before
                    R0data.R0ChargeRelaxation(j)=(Data.('Voltage (V)')(starts(3+(j-1)*3)-1)-Data.('Voltage (V)')(starts(3+(j-1)*3)))/abs(Data.('Current (A)')(starts(3+(j-1)*3)-1)); % R0 CR=(VG-VH)/Ic, where VG is the voltage before the start of the third rest and VH the voltage sampled immediately after
                    R0data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));

                    % Averages
                    R0data.R0DischargeAVG(j)=(R0data.R0Discharge(j)+R0data.R0DischargeRelaxation(j))/2;
                    R0data.R0ChargeAVG(j)=(R0data.R0Charge(j)+R0data.R0ChargeRelaxation(j))/2;
                end
                R0.(fields{i})=R0data;

                % Curve fitting to retrieve R1, C1, R2, C2
                R1data=table('Size',[numSteps 7],'VariableNames',["R1Discharge","R1DischargeRelaxation","R1Charge","R1ChargeRelaxation","R1DischargeAVG","R1ChargeAVG","SOC"],'VariableTypes',{'double','double','double','double','double','double','double'});
                C1data=table('Size',[numSteps 7],'VariableNames',["C1Discharge","C1DischargeRelaxation","C1Charge","C1ChargeRelaxation","C1DischargeAVG","C1ChargeAVG","SOC"],'VariableTypes',{'double','double','double','double','double','double','double'});
                R2data=table('Size',[numSteps 7],'VariableNames',["R2Discharge","R2DischargeRelaxation","R2Charge","R2ChargeRelaxation","R2DischargeAVG","R2ChargeAVG","SOC"],'VariableTypes',{'double','double','double','double','double','double','double'});
                C2data=table('Size',[numSteps 7],'VariableNames',["C2Discharge","C2DischargeRelaxation","C2Charge","C2ChargeRelaxation","C2DischargeAVG","C2ChargeAVG","SOC"],'VariableTypes',{'double','double','double','double','double','double','double'});
                adata=table('Size',[numSteps 5],'VariableNames',["aDischarge","aDischargeRelaxation","aCharge","aChargeRelaxation","SOC"],'VariableTypes',{'double','double','double','double','double'});
                bdata=table('Size',[numSteps 5],'VariableNames',["bDischarge","bDischargeRelaxation","bCharge","bChargeRelaxation","SOC"],'VariableTypes',{'double','double','double','double','double'});
                cdata=table('Size',[numSteps 5],'VariableNames',["cDischarge","cDischargeRelaxation","cCharge","cChargeRelaxation","SOC"],'VariableTypes',{'double','double','double','double','double'});
                tau1data=table('Size',[numSteps 5],'VariableNames',["tau1Discharge","tau1DischargeRelaxation","tau1Charge","tau1ChargeRelaxation","SOC"],'VariableTypes',{'double','double','double','double','double'});
                tau2data=table('Size',[numSteps 5],'VariableNames',["tau2Discharge","tau2DischargeRelaxation","tau2Charge","tau2ChargeRelaxation","SOC"],'VariableTypes',{'double','double','double','double','double'});
                V1data=table('Size',[numSteps 5],'VariableNames',["tau1Discharge","tau1DischargeRelaxation","tau1Charge","tau1ChargeRelaxation","SOC"],'VariableTypes',{'double','double','double','double','double'});
                V2data=table('Size',[numSteps 5],'VariableNames',["tau2Discharge","tau2DischargeRelaxation","tau2Charge","tau2ChargeRelaxation","SOC"],'VariableTypes',{'double','double','double','double','double'});

                for j=1:numSteps
                    R1data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    C1data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    R2data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    C2data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    adata.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    bdata.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    cdata.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    tau1data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    tau2data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    V1data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));
                    V2data.SOC(j)=Data.('SOC')(ends(1+(j-1)*3));

                    % Discharge
                    Dischargedata=Data(ends(1+(j-1)*3)+1:starts(2+(j-1)*3)-1,:);
                    DischargeArray{j}=Dischargedata;

                    Dischargedata.('Voltage (V)')=Dischargedata.('Voltage (V)')-Dischargedata.('Voltage (V)')(1); % V(t)-VOC+V0
                    Dischargedata.('Time (s)')=Dischargedata.('Time (s)')-Dischargedata.('Time (s)')(1); % Relative time
                    
                    ft=fittype('- a*(1 - exp(-x/tau1)) - b*(1 - exp(-x/tau2)) + c', ...
                                 'independent', 'x', ...
                                 'coefficients', {'a', 'tau1', 'b', 'tau2', 'c'});
                    opts = fitoptions(ft);
                    %[a, tau1, b, tau2, c]
                    opts.Lower=[1e-5, 0.05, 1e-5, 5, -1]; 
                    opts.Upper=[1.0, 5, 1.0, 200.0, +1];
                    opts.StartPoint=[0.05, 0.5, 0.05, 30.0, 0];
                    Dfit=fit(Dischargedata.('Time (s)'), Dischargedata.('Voltage (V)'), ft, opts);

                    R1data.R1Discharge(j)=Dfit.a/abs(Data.('Current (A)')(ends(1+(j-1)*3)+1)); % R1=a/Id
                    C1data.C1Discharge(j)=Dfit.tau1/R1data.R1Discharge(j); % C1=tau1/R1
                    R2data.R2Discharge(j)=Dfit.b/abs(Data.('Current (A)')(ends(1+(j-1)*3)+1)); % R2=b/Id
                    C2data.C2Discharge(j)=Dfit.tau2/R2data.R2Discharge(j); % C2=tau2/R2
                    adata.aDischarge(j)=Dfit.a;
                    bdata.bDischarge(j)=Dfit.b;
                    cdata.cDischarge(j)=Dfit.c;
                    tau1data.tau1Discharge(j)=Dfit.tau1;
                    tau2data.tau2Discharge(j)=Dfit.tau2;

                    % Discharge Relaxation
                    DischargeRelaxationdata=Data(starts(2+(j-1)*3):ends(2+(j-1)*3),:);
                    DischargeRelaxationArray{j}=DischargeRelaxationdata;
                    
                    InstantOCV=interp1(OCVdata.SOC, OCVdata.OCV, DischargeRelaxationdata.SOC(1), 'linear', 'extrap'); % OCV interpolation
                    DischargeRelaxationdata.('Voltage (V)')=DischargeRelaxationdata.('Voltage (V)')-InstantOCV; % V(t)-VOC
                    DischargeRelaxationdata.('Time (s)')=DischargeRelaxationdata.('Time (s)')-DischargeRelaxationdata.('Time (s)')(1); % Relative time
                    
                    ft = fittype('- a*exp(-x/tau1) - b*exp(-x/tau2) + c', ...
                                 'independent', 'x', ...
                                 'coefficients', {'a', 'tau1', 'b', 'tau2', 'c'});
                    opts = fitoptions(ft);
                    %[a, tau1, b, tau2, c]
                    opts.Lower=[1e-5, 0.05, 1e-5, 5, -1]; 
                    opts.Upper=[1.0, 5, 1.0, 200.0, +1];
                    opts.StartPoint=[0.05, 0.5, 0.05, 30.0, 0];
                    DRfit=fit(DischargeRelaxationdata.('Time (s)'), DischargeRelaxationdata.('Voltage (V)'), ft, opts);
                    
                    R1data.R1DischargeRelaxation(j)=DRfit.a/abs(Data.('Current (A)')(ends(1+(j-1)*3)+1))/(1-exp(-Dischargedata.('Time (s)')(end)/Dfit.tau1)); % R1=a/Id/(1-exp(-Td/tau1d)));
                    C1data.C1DischargeRelaxation(j)=DRfit.tau1/R1data.R1DischargeRelaxation(j); % C1=tau1/R1;
                    R2data.R2DischargeRelaxation(j)=DRfit.b/abs(Data.('Current (A)')(ends(1+(j-1)*3)+1))/(1-exp(-Dischargedata.('Time (s)')(end)/Dfit.tau2)); % R2=b/Id/(1-exp(-Td/tau2d)));
                    C2data.C2DischargeRelaxation(j)=DRfit.tau2/R2data.R2DischargeRelaxation(j); % C2=tau2/R2;
                    adata.aDischargeRelaxation(j)=DRfit.a;
                    bdata.bDischargeRelaxation(j)=DRfit.b;
                    cdata.cDischargeRelaxation(j)=DRfit.c;
                    tau1data.tau1DischargeRelaxation(j)=DRfit.tau1;
                    tau2data.tau2DischargeRelaxation(j)=DRfit.tau2;

                    % Charge
                    Chargedata=Data(ends(2+(j-1)*3)+1:starts(3+(j-1)*3)-1,:);
                    ChargeArray{j}=Chargedata;
                    V1_res=DRfit.a*exp(-DischargeRelaxationdata.('Time (s)')(end)/DRfit.tau1); % V1(0)=V1(0)dr*exp(-T_dr/tau1dr)
                    V2_res=DRfit.b*exp(-DischargeRelaxationdata.('Time (s)')(end)/DRfit.tau2); % V2(0)=V2(0)dr*exp(-T_dr/tau1dr)
                    
                    InstantOCV=interp1(OCVdata.SOC, OCVdata.OCV, Chargedata.SOC(1), 'linear', 'extrap'); % OCV interpolation
                    Chargedata.('Voltage (V)')=Chargedata.('Voltage (V)')-InstantOCV-R0data.R0Charge(j)*abs(Data.('Current (A)')(ends(2+(j-1)*3)+1)); % V(t)-VOC-V0
                    Chargedata.('Time (s)')=Chargedata.('Time (s)')-Chargedata.('Time (s)')(1); % Relative time
                    
                    ft=fittype('+ a*(1 - exp(-x/tau1)) + b*(1 - exp(-x/tau2)) - V1*exp(-x/tau1) - V2*exp(-x/tau2) +c', ...
                                 'independent', 'x', ...
                                 'coefficients', {'a', 'tau1', 'b', 'tau2', 'c'}, ...
                                 'problem', {'V1', 'V2'});
                    opts = fitoptions(ft);
                    %[a, tau1, b, tau2, c]
                    opts.Lower=[1e-5, 0.05, 1e-5, 5, -1]; 
                    opts.Upper=[1.0, 5, 1.0, 200.0, +1];
                    opts.StartPoint=[0.05, 0.5, 0.05, 30.0, 0];
                    Cfit=fit(Chargedata.('Time (s)'), Chargedata.('Voltage (V)'), ft, opts,'problem', {V1_res, V2_res});
                    
                    R1data.R1Charge(j)=Cfit.a/abs(Data.('Current (A)')(ends(2+(j-1)*3)+1)); % R1=a/Ic
                    C1data.C1Charge(j)=Cfit.tau1/R1data.R1Charge(j); % C1=tau1/R1
                    R2data.R2Charge(j)=Cfit.b/abs(Data.('Current (A)')(ends(2+(j-1)*3)+1)); % R2=b/Ic
                    C2data.C2Charge(j)=Cfit.tau2/R2data.R2Charge(j); % C2=tau2/R2
                    adata.aCharge(j)=Cfit.a;
                    bdata.bCharge(j)=Cfit.b;
                    cdata.cCharge(j)=Cfit.c;
                    tau1data.tau1Charge(j)=Cfit.tau1;
                    tau2data.tau2Charge(j)=Cfit.tau2;
                    V1data.V1Charge(j)=V1_res;
                    V2data.V2Charge(j)=V2_res;

                    % Charge Relaxation
                    ChargeRelaxationdata=Data(starts(3+(j-1)*3):ends(3+(j-1)*3),:);
                    ChargeRelaxationArray{j}=ChargeRelaxationdata;
                    V1_res_r=V1_res*exp(-Chargedata.('Time (s)')(end)/Cfit.tau1); % V1res(0)=V1c(0)*exp(-T_c/tau1c)
                    V2_res_r=V2_res*exp(-Chargedata.('Time (s)')(end)/Cfit.tau2); % V2res(0)=V2c(0)*exp(-T_c/tau2c)
                    
                    InstantOCV=interp1(OCVdata.SOC, OCVdata.OCV, ChargeRelaxationdata.SOC(1), 'linear', 'extrap'); % OCV interpolation
                    ChargeRelaxationdata.('Voltage (V)')=ChargeRelaxationdata.('Voltage (V)')-InstantOCV; % V(t)-VOC
                    ChargeRelaxationdata.('Time (s)')=ChargeRelaxationdata.('Time (s)')-ChargeRelaxationdata.('Time (s)')(1); % Relative time
                    
                    ft = fittype('+ a*exp(-x/tau1) + b*exp(-x/tau2) + c', ...
                                 'independent', 'x', ...
                                 'coefficients', {'a', 'tau1', 'b', 'tau2', 'c'});
                    opts = fitoptions(ft);
                    %[a, tau1, b, tau2]
                    opts.Lower=[1e-5, 0.05, 1e-5, 5, -1]; 
                    opts.Upper=[1.0, 5, 1.0, 200.0, +1];
                    opts.StartPoint=[0.05, 0.5, 0.05, 30.0, 0];
                    CRfit=fit(ChargeRelaxationdata.('Time (s)'), ChargeRelaxationdata.('Voltage (V)'), ft, opts);
                    
                    R1data.R1ChargeRelaxation(j)=(CRfit.a+V1_res_r)/abs(Data.('Current (A)')(ends(2+(j-1)*3)+1))/(1-exp(-Chargedata.('Time (s)')(end)/Cfit.tau1)); % R1=(a+V1res(0))/Ic/(1-exp(-Tc/tau1c)));
                    C1data.C1ChargeRelaxation(j)=CRfit.tau1/R1data.R1ChargeRelaxation(j); % C1=tau1/R1;
                    R2data.R2ChargeRelaxation(j)=(CRfit.b+V2_res_r)/abs(Data.('Current (A)')(ends(2+(j-1)*3)+1))/(1-exp(-Chargedata.('Time (s)')(end)/Cfit.tau2)); % R2=(b+V2res(0))/Ic/(1-exp(-Tc/tau2c)));
                    C2data.C2ChargeRelaxation(j)=CRfit.tau2/R2data.R2ChargeRelaxation(j); % C2=tau2/R2;
                    adata.aChargeRelaxation(j)=CRfit.a;
                    bdata.bChargeRelaxation(j)=CRfit.b;
                    cdata.cChargeRelaxation(j)=CRfit.c;
                    tau1data.tau1ChargeRelaxation(j)=CRfit.tau1;
                    tau2data.tau2ChargeRelaxation(j)=CRfit.tau2;
                    
                    % Averages
                    R1data.R1ChargeAVG(j)=(R1data.R1Charge(j)+R1data.R1ChargeRelaxation(j))/2;
                    R1data.R1DischargeAVG(j)=(R1data.R1Discharge(j)+R1data.R1DischargeRelaxation(j))/2;

                    C1data.C1ChargeAVG(j)=(C1data.C1Charge(j)+C1data.C1ChargeRelaxation(j))/2;
                    C1data.C1DischargeAVG(j)=(C1data.C1Discharge(j)+C1data.C1DischargeRelaxation(j))/2;

                    R2data.R2ChargeAVG(j)=(R2data.R2Charge(j)+R2data.R2ChargeRelaxation(j))/2;
                    R2data.R2DischargeAVG(j)=(R2data.R2Discharge(j)+R2data.R2DischargeRelaxation(j))/2;

                    C2data.C2ChargeAVG(j)=(C2data.C2Charge(j)+C2data.C2ChargeRelaxation(j))/2;
                    C2data.C2DischargeAVG(j)=(C2data.C2Discharge(j)+C2data.C2DischargeRelaxation(j))/2;
                end
                R1.(fields{i})=R1data;
                C1.(fields{i})=C1data;
                R2.(fields{i})=R2data;
                C2.(fields{i})=C2data;
                Discharge.(fields{i})=DischargeArray;
                DischargeRelaxation.(fields{i})=DischargeRelaxationArray;
                Charge.(fields{i})=ChargeArray;
                ChargeRelaxation.(fields{i})=ChargeRelaxationArray;
                a.(fields{i})=adata;
                b.(fields{i})=bdata;
                c.(fields{i})=cdata;
                tau1.(fields{i})=tau1data;
                tau2.(fields{i})=tau2data;
                V1.(fields{i})=V1data;
                V2.(fields{i})=V2data;

                % Plot OCV
                figure;
                %clf;
                plot(OCVdata.SOC, OCVdata.OCV, 'LineWidth', 2, 'DisplayName', 'OCV');
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('OCV (V)');
                title(['OCV vs SOC - Test',fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot R0
                figure;
                %clf;
                hold on;
                plot(R0data.SOC, R0data.R0Discharge, 'LineWidth', 2, 'DisplayName', 'R0 Discharge');
                plot(R0data.SOC, R0data.R0DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R0 Discharge Relaxation');
                %plot(R0data.SOC, R0data.R0DischargeAVG, 'LineWidth', 2, 'DisplayName', 'R0 Discharge AVERAGE');
                plot(R0data.SOC, R0data.R0Charge, 'LineWidth', 2, 'DisplayName', 'R0 Charge');
                plot(R0data.SOC, R0data.R0ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R0 Charge Relaxation');
                %plot(R0data.SOC, R0data.R0ChargeAVG, 'LineWidth', 2, 'DisplayName', 'R0 Charge AVERAGE');
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('R0 (Ohm)');
                title(['R0 vs SOC - Test',fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot R1
                figure;
                %clf;
                hold on;
                plot(R1data.SOC, R1data.R1Discharge, 'LineWidth', 2, 'DisplayName', 'R1 Discharge');
                plot(R1data.SOC, R1data.R1DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R1 Discharge Relaxation');
                %plot(R1data.SOC, R1data.R1DischargeAVG, 'LineWidth', 2, 'DisplayName', 'R1 Discharge AVERAGE');
                plot(R1data.SOC, R1data.R1Charge, 'LineWidth', 2, 'DisplayName', 'R1 Charge');
                plot(R1data.SOC, R1data.R1ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R1 Charge Relaxation');
                %plot(R1data.SOC, R1data.R1ChargeAVG, 'LineWidth', 2, 'DisplayName', 'R1 Charge AVERAGE');
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('R1 (Ohm)');
                title(['R1 vs SOC - Test',fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot C1
                figure;
                %clf;
                hold on;
                plot(C1data.SOC, C1data.C1Discharge, 'LineWidth', 2, 'DisplayName', 'C1 Discharge');
                plot(C1data.SOC, C1data.C1DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'C1 Discharge Relaxation');
                %plot(C1data.SOC, C1data.C1DischargeAVG, 'LineWidth', 2, 'DisplayName', 'C1 Discharge AVERAGE');
                plot(C1data.SOC, C1data.C1Charge, 'LineWidth', 2, 'DisplayName', 'C1 Charge');
                plot(C1data.SOC, C1data.C1ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'C1 Charge Relaxation');
                %plot(C1data.SOC, C1data.C1ChargeAVG, 'LineWidth', 2, 'DisplayName', 'C1 Charge AVERAGE');
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('C1 (F)');
                title(['C1 vs SOC - Test',fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot R2
                figure;
                %clf;
                hold on;
                plot(R2data.SOC, R2data.R2Discharge, 'LineWidth', 2, 'DisplayName', 'R2 Discharge');
                plot(R2data.SOC, R2data.R2DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R2 Discharge Relaxation');
                %plot(R2data.SOC, R2data.R2DischargeAVG, 'LineWidth', 2, 'DisplayName', 'R2 Discharge AVERAGE');
                plot(R2data.SOC, R2data.R2Charge, 'LineWidth', 2, 'DisplayName', 'R2 Charge');
                plot(R2data.SOC, R2data.R2ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R2 Charge Relaxation');
                %plot(R2data.SOC, R2data.R2ChargeAVG, 'LineWidth', 2, 'DisplayName', 'R2 Charge AVERAGE');
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('R2 (Ohm)');
                title(['R2 vs SOC - Test',fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot C2
                figure;
                %clf;
                hold on;
                plot(C2data.SOC, C2data.C2Discharge, 'LineWidth', 2, 'DisplayName', 'C2 Discharge');
                plot(C2data.SOC, C2data.C2DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'C2 Discharge Relaxation');
                %plot(C2data.SOC, C2data.C2DischargeAVG, 'LineWidth', 2, 'DisplayName', 'C2 Discharge AVERAGE');
                plot(C2data.SOC, C2data.C2Charge, 'LineWidth', 2, 'DisplayName', 'C2 Charge');
                plot(C2data.SOC, C2data.C2ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'C2 Charge Relaxation');
                %plot(C2data.SOC, C2data.C2ChargeAVG, 'LineWidth', 2, 'DisplayName', 'C2 Charge AVERAGE');
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('C2 (F)');
                title(['C2 vs SOC - Test',fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Display data
                fprintf('Results manual method (fittype):');
                disp(R0data);
                disp(R1data);
                disp(C1data);
                disp(R2data);
                disp(C2data);

                % Plot voltage different phases
                figure;
                hold on;
                plota=plot(Data.('Time (s)'), Data.('Voltage (V)') ,'-b', 'LineWidth', 2, 'DisplayName', 'All');
                for j=1:numSteps
                    plotd=plot(DischargeArray{j}.('Time (s)'), DischargeArray{j}.('Voltage (V)') ,'-r', 'LineWidth', 2);
                    % Verification plot Discharge
                    Vocv=OCVdata.OCV(j);
                    R0temp=R0data.R0Discharge(j);
                    Id=abs(Data.('Current (A)')(ends(1+(j-1)*3)+1));
                    R1temp=R1data.R1Discharge(j);
                    R2temp=R2data.R2Discharge(j);
                    t=DischargeArray{j}.('Time (s)')-DischargeArray{j}.('Time (s)')(1);
                    tau1temp=tau1data.tau1Discharge(j);
                    tau2temp=tau2data.tau2Discharge(j);
                    %atemp=adata.aDischarge(j);
                    %btemp=bdata.bDischarge(j);
                    ctemp=cdata.cDischarge(j);
                    EstimatedVoltage = Vocv+ctemp-R0temp.*Id-R1temp.*Id.*(1-exp(-t./tau1temp))-R2temp.*Id.*(1-exp(-t./tau2temp));
                    Vplot=plot(DischargeArray{j}.('Time (s)'), EstimatedVoltage, '--k', 'LineWidth', 2);
                    
                    plotdr=plot(DischargeRelaxationArray{j}.('Time (s)'), DischargeRelaxationArray{j}.('Voltage (V)') ,'-g', 'LineWidth', 2);
                    % Verification plot Discharge Relaxation
                    Vocv=interp1(OCVdata.SOC, OCVdata.OCV, DischargeRelaxationArray{j}.SOC(1), 'linear', 'extrap');
                    %R0temp=R0data.R0DischargeRelaxation(j);
                    Id=abs(Data.('Current (A)')(ends(1+(j-1)*3)+1));
                    R1temp=R1data.R1DischargeRelaxation(j);
                    R2temp=R2data.R2DischargeRelaxation(j);
                    t=DischargeRelaxationArray{j}.('Time (s)')-DischargeRelaxationArray{j}.('Time (s)')(1);
                    Td=DischargeArray{j}.('Time (s)')(end)-DischargeArray{j}.('Time (s)')(1);
                    tau1temp=tau1data.tau1DischargeRelaxation(j);
                    tau1prev=tau1data.tau1Discharge(j);
                    tau2temp=tau2data.tau2DischargeRelaxation(j);
                    tau2prev=tau2data.tau2Discharge(j);
                    %atemp=adata.aDischargeRelaxation(j);
                    %btemp=bdata.bDischargeRelaxation(j);
                    ctemp=cdata.cDischargeRelaxation(j);
                    EstimatedVoltage = Vocv+ctemp-R1temp.*Id.*(1-exp(-Td/tau1prev)).*exp(-t./tau1temp)-R2temp.*Id.*(1-exp(-Td/tau2prev)).*exp(-t./tau2temp);
                    plot(DischargeRelaxationArray{j}.('Time (s)'), EstimatedVoltage, '--k', 'LineWidth', 2);

                    plotc=plot(ChargeArray{j}.('Time (s)'), ChargeArray{j}.('Voltage (V)') ,'-c', 'LineWidth', 2);
                    % Verification plot Charge
                    Vocv=interp1(OCVdata.SOC, OCVdata.OCV, ChargeArray{j}.SOC(1), 'linear', 'extrap');
                    R0temp=R0data.R0Charge(j);
                    Ic=abs(Data.('Current (A)')(ends(2+(j-1)*3)+1));
                    R1temp=R1data.R1Charge(j);
                    R2temp=R2data.R2Charge(j);
                    t=ChargeArray{j}.('Time (s)')-ChargeArray{j}.('Time (s)')(1);
                    tau1temp=tau1data.tau1Charge(j);
                    tau2temp=tau2data.tau2Charge(j);
                    V1temp=V1data.V1Charge(j);
                    V2temp=V2data.V2Charge(j);
                    %atemp=adata.aCharge(j);
                    %btemp=bdata.bCharge(j);
                    ctemp=cdata.cCharge(j);
                    EstimatedVoltage = Vocv+ctemp+R0temp.*Ic+R1temp.*Ic.*(1-exp(-t./tau1temp))+R2temp.*Ic.*(1-exp(-t./tau2temp))-V1temp.*exp(-t/tau1temp)-V2temp.*exp(-t/tau2temp);
                    plot(ChargeArray{j}.('Time (s)'), EstimatedVoltage, '--k', 'LineWidth', 2);

                    plotcr=plot(ChargeRelaxationArray{j}.('Time (s)'), ChargeRelaxationArray{j}.('Voltage (V)') ,'-m', 'LineWidth', 2);
                    % Verification plot Charge Relaxation
                    Vocv=interp1(OCVdata.SOC, OCVdata.OCV, ChargeRelaxationArray{j}.SOC(1), 'linear', 'extrap');
                    %R0temp=R0data.R0ChargeRelaxation(j);
                    Ic=abs(Data.('Current (A)')(ends(2+(j-1)*3)+1));
                    R1temp=R1data.R1ChargeRelaxation(j);
                    R2temp=R2data.R2ChargeRelaxation(j);
                    t=ChargeRelaxationArray{j}.('Time (s)')-ChargeRelaxationArray{j}.('Time (s)')(1);
                    Tc=ChargeArray{j}.('Time (s)')(end)-ChargeArray{j}.('Time (s)')(1);
                    tau1temp=tau1data.tau1ChargeRelaxation(j);
                    tau1prev=tau1data.tau1Charge(j);
                    tau2temp=tau2data.tau2ChargeRelaxation(j);
                    tau2prev=tau2data.tau2Charge(j);
                    V1temp=V1data.V1Charge(j);
                    V2temp=V2data.V2Charge(j);
                    %atemp=adata.aChargeRelaxation(j);
                    %btemp=bdata.bChargeRelaxation(j);
                    ctemp=cdata.cChargeRelaxation(j);
                    EstimatedVoltage = Vocv+ctemp+R1temp.*Ic.*(1-exp(-Tc/tau1prev)).*exp(-t./tau1temp)-V1temp.*exp(-Tc./tau1prev).*exp(-t./tau1temp)+R2temp.*Ic.*(1-exp(-Tc/tau2prev)).*exp(-t./tau2temp)-V2temp.*exp(-Tc./tau2prev).*exp(-t./tau2temp);
                    plot(ChargeRelaxationArray{j}.('Time (s)'), EstimatedVoltage, '--k', 'LineWidth', 2);
                end
                hold off;
                xlabel('Time (s)');
                ylabel('Voltage (V)');
                title(['Voltage vs Time - Test',fields{i}], 'Interpreter','none');
                legend([plota, plotd, plotdr, plotc, plotcr, Vplot], {'All','Discharge','Discharge Relaxation','Charge','Charge Relaxation','Interpolated Data'}, 'location', 'best')
                grid on;
                set(gcf,'Color','white');

                % Plot current different phases
                figure;
                hold on;
                plota=plot(Data.('Time (s)'), Data.('Current (A)') ,'-b', 'LineWidth', 2, 'DisplayName', 'All');
                for j=1:numSteps
                    plotd=plot(DischargeArray{j}.('Time (s)'), DischargeArray{j}.('Current (A)') ,'-r', 'LineWidth', 2); % Discharge 
                    plotdr=plot(DischargeRelaxationArray{j}.('Time (s)'), DischargeRelaxationArray{j}.('Current (A)') ,'-g', 'LineWidth', 2); % Discharge relaxation
                    plotc=plot(ChargeArray{j}.('Time (s)'), ChargeArray{j}.('Current (A)') ,'-c', 'LineWidth', 2); % Charge
                    plotcr=plot(ChargeRelaxationArray{j}.('Time (s)'), ChargeRelaxationArray{j}.('Current (A)') ,'-m', 'LineWidth', 2); % Charge relaxation
                end
                hold off;
                xlabel('Time (s)');
                ylabel('Current (A)');
                title(['Current vs Time - Test',fields{i}], 'Interpreter','none');
                legend([plota, plotd, plotdr, plotc, plotcr], {'All','Discharge','Discharge Relaxation','Charge','Charge Relaxation'}, 'location', 'best')
                grid on;
                set(gcf,'Color','white');
            end

           case 10
            %% EXPORT TO EXCEL FILE (MANUAL CODE RESULTS)
            %% MATLAB Code to Export Data to Excel
            % Output filename
            filename = 'Exported_Battery_Params_MANUAL.xlsx';
            
            % Define the exact column headers
            headers = { ...
                'SOC', 'OCV', ...
                'R0 D', 'R0 DR', 'R0 DA', 'R0 C', 'R0 CR', 'R0 CA', ...
                'R1 D', 'R1 DR', 'R1 DA', 'R1 C', 'R1 CR', 'R1 CA', ...
                'C1 D', 'C1 DR', 'C1 DA', 'C1 C', 'C1 CR', 'C1 CA', ...
                'R2 D', 'R2 DR', 'R2 DA', 'R2 C', 'R2 CR', 'R2 CA', ...
                'C2 D', 'C2 DR', 'C2 DA', 'C2 C', 'C2 CR', 'C2 CA' ...
            };
            
            % Helper function to ensure data is a column vector
            toCol = @(x) x(:);
            
            % Extract and Concatenate Data
            dataMatrix = [ ...
                toCol(OCVdata.SOC), ...                                         % SOC
                toCol(OCVdata.OCV), ...                                         % OCV
                toCol(R0data.R0Discharge), toCol(R0data.R0DischargeRelaxation), toCol(R0data.R0DischargeAVG), toCol(R0data.R0Charge), toCol(R0data.R0ChargeRelaxation), toCol(R0data.R0ChargeAVG), ... % R0
                toCol(R1data.R1Discharge), toCol(R1data.R1DischargeRelaxation), toCol(R1data.R1DischargeAVG), toCol(R1data.R1Charge), toCol(R1data.R1ChargeRelaxation), toCol(R1data.R1ChargeAVG), ... % R1
                toCol(C1data.C1Discharge), toCol(C1data.C1DischargeRelaxation), toCol(C1data.C1DischargeAVG), ... % C1 (Discharge parts)
                toCol(C1data.C1Charge),    toCol(C1data.C1ChargeRelaxation),    toCol(C1data.C1ChargeAVG), ...     % C1 (Charge parts)
                toCol(R2data.R2Discharge), toCol(R2data.R2DischargeRelaxation), toCol(R2data.R2DischargeAVG), toCol(R2data.R2Charge), toCol(R2data.R2ChargeRelaxation), toCol(R2data.R2ChargeAVG), ... % R2
                toCol(C2data.C2Discharge), toCol(C2data.C2DischargeRelaxation), toCol(C2data.C2DischargeAVG), toCol(C2data.C2Charge), toCol(C2data.C2ChargeRelaxation), toCol(C2data.C2ChargeAVG)      % C2
            ];
            
            % Prepare Cell Array for Export
            dataCell = num2cell(dataMatrix);
            
            % Combine headers and data
            outputBlock = [headers; dataCell];
            
            % Write to Excel
            try
                writecell(outputBlock, filename, 'Sheet', 1, 'Range', 'A2');
                writecell({'Manual (fittype)'}, filename, 'Sheet', 1, 'Range', 'A1');
                disp(['Successfully exported data to ', filename]);
            catch ME
                error('Error writing to file. Make sure the Excel file is not open. \n%s', ME.message);
            end

           case 11
            %% CREATE HPPC DATA (MATLAB AUTO CODE)
            % Initialize the combined struct (merge of NEWARE and DEWESOFT Data)
            DataCOMBINED = struct();
            DataHPPC = struct();

            % Get the field names of both structs
            fieldsNEWARE = fieldnames(DataNEWARE);
            fieldsDEWESOFT = fieldnames(DataDEWESOFT);

            % Loop over the fields of DataNEWARE to find corresponding fields in DataDEWESOFT
            for i = 1:length(fieldsNEWARE)
                % Check if the current field name ends with 'NEWARE'
                if endsWith(fieldsNEWARE{i}, '_NEWARE_xlsx')
                    % Extract the base name by removing 'NEWARE'
                    baseName = extractBefore(fieldsNEWARE{i}, '_NEWARE_xlsx');
                    % Check if the same base name exists in DataDEWESOFT with 'DEWESOFT'
                    correspondingField = strcat(baseName, '_DEWESOFT_xlsx');
                    if isfield(DataDEWESOFT, correspondingField)
                        % Get the tables from both structs
                        tableNEWARE = DataNEWARE.(fieldsNEWARE{i});
                        tableDEWESOFT = DataDEWESOFT.(correspondingField);

                        % Rename 'Time' to 'Time(s)' for DEWESOFT
                        tableDEWESOFT.Properties.VariableNames{strcmp(tableDEWESOFT.Properties.VariableNames, 'Time')} = 'Time(s)';
                        
                        % Merge the tables using an outer join on 'Time(s)'
                        combinedTable = outerjoin(tableNEWARE, tableDEWESOFT, 'Keys', 'Time(s)', 'MergeKeys', true);

                        % Fill Nan and empty strings with data
                        filledTable = combinedTable;
                        varNames = filledTable.Properties.VariableNames;
                    
                        for k = 1:numel(varNames)
                            colData = filledTable.(varNames{k});
                    
                            % Handle numeric columns
                            if isnumeric(colData)
                                % Forward fill
                                lastVal = NaN;
                                for j = 1:length(colData)
                                    if ~isnan(colData(j))
                                        lastVal = colData(j);
                                    elseif ~isnan(lastVal)
                                        colData(j) = lastVal;
                                    end
                                end
                                % Fill top NaNs with first valid value
                                if isnan(colData(1))
                                    firstValid = find(~isnan(colData), 1, 'first');
                                    if ~isempty(firstValid)
                                        colData(1:firstValid-1) = colData(firstValid);
                                    end
                                end
                                filledTable.(varNames{k}) = colData;
                    
                            % Handle string/cell string columns
                            elseif iscellstr(colData) || isstring(colData)
                                % Convert to cellstr if it's string for uniform handling
                                if isstring(colData)
                                    colData = cellstr(colData);
                                    convertBackToString = true;
                                else
                                    convertBackToString = false;
                                end
                    
                                lastStr = '';
                                for j = 1:length(colData)
                                    if ~strcmp(colData{j}, '')
                                        lastStr = colData{j};
                                    elseif ~strcmp(lastStr, '')
                                        colData{j} = lastStr;
                                    end
                                end
                    
                                % Fill top empty strings with first non-empty
                                if strcmp(colData{1}, '')
                                    firstValid = find(~cellfun(@isempty, colData), 1, 'first');
                                    if ~isempty(firstValid)
                                        colData(1:firstValid-1) = repmat(colData(firstValid), firstValid-1, 1);
                                    end
                                end
                    
                                % Convert back if necessary
                                if convertBackToString
                                    colData = string(colData);
                                end
                    
                                filledTable.(varNames{k}) = colData;
                            end
                        end

                        % Keep only HPPC test (delete inititial charge/discharge) by starting to consider data from the 4th 'Rest' phase
                        is_rest = strcmp(filledTable.('Step Type'), 'Rest'); % Logical index of 'Rest' rows
                        transitions = diff([0; is_rest; 0]);  % Find the start indices of all 'Rest' blocks
                        starts = find(transitions == 1);      % start of each block
                        ends   = find(transitions == -1) - 1; % end of each block (not used here)
                        if length(starts) < 4 % Check if there are at least 4 'Rest' blocks
                            disp('Less than 4 ''Rest'' blocks found.');
                            trimmedTable = table();  % return empty table
                        else
                            start_idx = starts(4); % Get the start index of the 4th block
                            trimmedTable = filledTable(start_idx:end, :); % Keep the table from that point onward
                            trimmedTable.('Time(s)') = trimmedTable.('Time(s)')-trimmedTable.('Time(s)')(1); % Adjust time
                        end

                        % Keep only HPPC test (delete final discharge) by finishing to consider up to the second to last 'Rest' phase
                        is_rest = strcmp(trimmedTable.('Step Type'), 'Rest'); % Logical index of 'Rest' rows
                        transitions = diff([0; is_rest; 0]); % Identify block starts and ends
                        starts = find(transitions == 1);
                        ends   = find(transitions == -1) - 1;
                        if length(starts) < 2 % Ensure there are at least 2 blocks
                            disp('Less than two ''Rest'' blocks found.');
                            trimmedTable2 = trimmedTable;  % return original table
                        else
                            cut_idx = ends(end-1);  % Get the end index of the second-to-last block  
                            trimmedTable2 = trimmedTable(1:cut_idx, :); % Keep only rows up to that point
                        end

                        % Save the combined table in the new struct
                        DataCOMBINED.(baseName) = trimmedTable2;

                        % Save the combined table in the new struct (Only useful data)
                        DataHPPC.(baseName) = DataCOMBINED.(baseName)(:, {'Step Type','Time(s)', 'Current(A)', 'Voltage(V)', 'Temperature (C)','SOC'});
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Time(s)')} = 'Time (s)';
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Current(A)')} = 'Current (A)';
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Voltage(V)')} = 'Voltage (V)';
                    end
                end
            end

            fields=fieldnames(DataHPPC);
            HPPCAuto=struct();
            
            for i = 1:length(fields) % Iterating all the tests
                Data=DataHPPC.(fields{i});
                hppcDataInput = Data(:, ["Time (s)", "Voltage (V)", "Current (A)"]);

                % Using Matlab Integrated HPPC Tool (only from R2025)
                hppcResult = hppcTest(hppcDataInput,...
                    TimeVariable="Time (s)", ...
                    VoltageVariable="Voltage (V)", ...
                    CurrentVariable="Current (A)", ...
                    TemperatureVariable="Temperature (C)", ...
                    StateofChargeVariable="SOC", ...
                    Capacity=DataNEWARE.(strcat(fields{i},'_NEWARE_xlsx')).('Pilot discharge capacity (Ah)')(1), ...
                    InitialSOC=1);

                % Remove Unwanted Pulses at SOC=1
                removePulse(hppcResult,11);
                removePulse(hppcResult,1);
                
                % Save in Struct
                HPPCAuto.(fields{i})=hppcResult;

                % Display Summary
                disp(hppcResult.TestSummary)

                % Plot
                figure;
                plot(hppcResult);
                title(['Voltage vs Time - Test',fields{i}], 'Interpreter','none');
                grid on;
                set(gcf,'Color','white');

                % Fit parameters
                batteryEcmFM = fitECM(hppcResult, SegmentToFit="loadAndRelaxation", FittingMethod="fminsearch");
                batteryEcmCF = fitECM(hppcResult, SegmentToFit="loadAndRelaxation", FittingMethod="curvefit");

                % Display parameters
                fprintf('Results auto method (fminsearch):');
                disp(batteryEcmFM.ModelParameterTables);
                fprintf('Results auto method (curvefit):');
                disp(batteryEcmCF.ModelParameterTables);

                %Plotting parameters
                SOCvector=[0 10 20 30 40 50 60 70 80 90 100];

                % Plot OCV
                figure;
                %clf;
                hold on;
                plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeOpenCircuitVoltage, 'LineWidth', 2, 'DisplayName', 'OCV Discharge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeOpenCircuitVoltage, 'LineWidth', 2, 'DisplayName', 'OCV Discharge CURVEFIT');
                plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeOpenCircuitVoltage, 'LineWidth', 2, 'DisplayName', 'OCV Charge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeOpenCircuitVoltage, 'LineWidth', 2, 'DisplayName', 'OCV Charge CURVEFIT'); 
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('OCV (V)');
                title(['OCV vs SOC - Test',fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');
                
                % Plot R0
                figure;
                %clf;
                hold on;
                plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeR0, 'LineWidth', 2, 'DisplayName', 'R0 Discharge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeR0, 'LineWidth', 2, 'DisplayName', 'R0 Discharge CURVEFIT');
                plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeR0, 'LineWidth', 2, 'DisplayName', 'R0 Charge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeR0, 'LineWidth', 2, 'DisplayName', 'R0 Charge CURVEFIT'); 
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('R0 (Ohm)');
                title(['R0 vs SOC - Test',fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot R1
                figure;
                hold on;
                plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeR1, 'LineWidth', 2, 'DisplayName', 'R1 Discharge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeR1, 'LineWidth', 2, 'DisplayName', 'R1 Discharge CURVEFIT');
                plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeR1, 'LineWidth', 2, 'DisplayName', 'R1 Charge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeR1, 'LineWidth', 2, 'DisplayName', 'R1 Charge CURVEFIT'); 
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('R1 (Ohm)');
                title(['R1 vs SOC - Test ', fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot C1
                figure;
                hold on;
                plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeC1, 'LineWidth', 2, 'DisplayName', 'C1 Discharge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeC1, 'LineWidth', 2, 'DisplayName', 'C1 Discharge CURVEFIT');
                plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeC1, 'LineWidth', 2, 'DisplayName', 'C1 Charge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeC1, 'LineWidth', 2, 'DisplayName', 'C1 Charge CURVEFIT'); 
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('C1 (Farad)');
                title(['C1 vs SOC - Test ', fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot R2
                figure;
                hold on;
                plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeR2, 'LineWidth', 2, 'DisplayName', 'R2 Discharge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeR2, 'LineWidth', 2, 'DisplayName', 'R2 Discharge CURVEFIT');
                plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeR2, 'LineWidth', 2, 'DisplayName', 'R2 Charge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeR2, 'LineWidth', 2, 'DisplayName', 'R2 Charge CURVEFIT'); 
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('R2 (Ohm)');
                title(['R2 vs SOC - Test ', fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot C2
                figure;
                hold on;
                plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeC2, 'LineWidth', 2, 'DisplayName', 'C2 Discharge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeC2, 'LineWidth', 2, 'DisplayName', 'C2 Discharge CURVEFIT');
                plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeC2, 'LineWidth', 2, 'DisplayName', 'C2 Charge FMINSEARCH');
                plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeC2, 'LineWidth', 2, 'DisplayName', 'C2 Charge CURVEFIT'); 
                hold off;
                xlabel('SOC (%)');
                xlim([10 90]);
                ylabel('C2 (Farad)');
                title(['C2 vs SOC - Test ', fields{i}], 'Interpreter','none');
                legend('Location','best');
                grid on;
                set(gcf,'Color','white');

                % Plot time response
                figure;
                plot(batteryEcmFM,1:18);
                %title(['Pulse fit FMINSEARCH - Test', fields{i}], 'Interpreter','none');
                %grid on;
                set(gcf,'Color','white');

                figure;
                plot(batteryEcmCF,1:18);
                %title(['Pulse fit CURVEFIT - Test', fields{i}], 'Interpreter','none');
                %grid on;
                set(gcf,'Color','white');

                % Plot HPPC test simulation
                figure;
                simulateHPPCTest(batteryEcmFM,hppcResult);
                %title(['Simulate HPPC FMINSEARCH - Test', fields{i}], 'Interpreter','none');
                %grid on;
                set(gcf,'Color','white');

                figure;
                simulateHPPCTest(batteryEcmFM,hppcResult);
                %title(['Simulate HPPC CURVEFIT - Test', fields{i}], 'Interpreter','none');
                %grid on;
                set(gcf,'Color','white');
            end

           case 12
            %% EXPORT TO EXCEL FILE (AUTO CODE RESULTS)
            %FMINSEARCH
            % Check for input structure
            if exist('batteryEcmFM', 'var')
                T = batteryEcmFM.TestParameterTables;
            else
                error('Variable "batteryEcmFM" not found in workspace.');
            end
            
            % Output filename
            filename = 'Exported_Battery_Params_AUTO.xlsx';
            
            % Extract and Flip Data Vectors
            % Helper: Force to column vector (:), then flip up-down (flipud)
            toCol = @(x) flipud(x(:));
            
            % SOC (Manually defined inverse vector as in your snippet)
            % Note: Ensure the length matches your data (9 points here)
            SOCvectorinv = [90; 80; 70; 60; 50; 40; 30; 20; 10]; 
            
            % R0
            R0_D = toCol(T.DischargeR0);
            R0_C = toCol(T.ChargeR0);
            
            % R1
            R1_D = toCol(T.DischargeR1);
            R1_C = toCol(T.ChargeR1);
            
            % C1
            C1_D = toCol(T.DischargeC1);
            C1_C = toCol(T.ChargeC1);
            
            % OCV
            OCV_D = toCol(T.DischargeOpenCircuitVoltage);
            OCV_C = toCol(T.ChargeOpenCircuitVoltage);
            
            % R2
            R2_D = toCol(T.DischargeR2);
            R2_C = toCol(T.ChargeR2);
            
            % C2
            C2_D = toCol(T.DischargeC2);
            C2_C = toCol(T.ChargeC2);
            
            % Construct Data Matrix & Headers
            headers = { ...
                'SOC', ...
                'OCV D', 'OCV C', ...
                'R0 D', 'R0 C', ...
                'R1 D', 'R1 C', ...
                'C1 D', 'C1 C', ...
                'R2 D', 'R2 C', ...
                'C2 D', 'C2 C' ...
            };
            
            % Combine vectors into a matrix
            dataMatrix = [ ...
                SOCvectorinv, ...
                OCV_D, OCV_C, ...
                R0_D, R0_C, ...
                R1_D, R1_C, ...
                C1_D, C1_C, ...
                R2_D, R2_C, ...
                C2_D, C2_C ...
            ];
            
            % Write to Excel
            % Convert to cell array for writing with headers
            outputBlock = [headers; num2cell(dataMatrix)];
            
            try
                writecell({'Fminsearch'}, filename, 'Sheet', 1, 'Range', 'A1');
                writecell(outputBlock, filename, 'Sheet', 1, 'Range', 'A2');
                disp(['Successfully exported FM data to ', filename]);
            catch ME
                error('Error writing to file. Please ensure the Excel file is closed. \n%s', ME.message);
            end

            % Check for input structure
            if exist('batteryEcmCF', 'var')
                T = batteryEcmCF.TestParameterTables;
            else
                error('Variable "batteryEcmCF" not found in workspace.');
            end
            
            % Extract and Flip Data Vectors
            % Helper: Force to column vector (:), then flip up-down (flipud)
            toCol = @(x) flipud(x(:));
            
            % SOC
            SOCvectorinv=[90; 80; 70; 60; 50; 40; 30; 20; 10];

            % R0
            R0_D = toCol(T.DischargeR0);
            R0_C = toCol(T.ChargeR0);
            
            % R1
            R1_D = toCol(T.DischargeR1);
            R1_C = toCol(T.ChargeR1);
            
            % C1
            C1_D = toCol(T.DischargeC1);
            C1_C = toCol(T.ChargeC1);
            
            % OCV
            OCV_D = toCol(T.DischargeOpenCircuitVoltage);
            OCV_C = toCol(T.ChargeOpenCircuitVoltage);
            
            % R2
            R2_D = toCol(T.DischargeR2);
            R2_C = toCol(T.ChargeR2);
            
            % C2
            C2_D = toCol(T.DischargeC2);
            C2_C = toCol(T.ChargeC2);
            
            % Construct Data Matrix & Headers
            headers = { ...
                'SOC', ...
                'OCV D', 'OCV C', ...
                'R0 D', 'R0 C', ...
                'R1 D', 'R1 C', ...
                'C1 D', 'C1 C', ...
                'R2 D', 'R2 C', ...
                'C2 D', 'C2 C' ...
            };
            
            % Combine vectors into a matrix
            dataMatrix = [ ...
                SOCvectorinv, ...
                OCV_D, OCV_C, ...
                R0_D, R0_C, ...
                R1_D, R1_C, ...
                C1_D, C1_C, ...
                R2_D, R2_C, ...
                C2_D, C2_C ...
            ];
            
            % Write to Excel
            % Convert to cell array for writing with headers
            outputBlock = [headers; num2cell(dataMatrix)];
            
            try
                writecell({'Curvefit'}, filename, 'Sheet', 2, 'Range', 'A1');
                writecell(outputBlock, filename, 'Sheet', 2, 'Range', 'A2');
                disp(['Successfully exported CF data to ', filename]);
            catch ME
                error('Error writing to file. Please ensure the Excel file is closed. \n%s', ME.message);
            end

           case 13
            %% COMPARE HPPC RESULTS
            % Plot OCV
            figure;
            %clf;
            hold on;
            % Manual
            plot(OCVdata.SOC, OCVdata.OCV, 'LineWidth', 2, 'DisplayName', 'OCV');
            % Auto
            plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeOpenCircuitVoltage, 'LineWidth', 2, 'DisplayName', 'OCV Discharge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeOpenCircuitVoltage, 'LineWidth', 2, 'DisplayName', 'OCV Discharge CURVEFIT');
            plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeOpenCircuitVoltage, 'LineWidth', 2, 'DisplayName', 'OCV Charge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeOpenCircuitVoltage, 'LineWidth', 2, 'DisplayName', 'OCV Charge CURVEFIT'); 
            hold off;
            xlabel('SOC (%)');
            xlim([10 90]);
            ylabel('OCV (V)');
            title(['OCV vs SOC - Test',fields{i}], 'Interpreter','none');
            legend('Location','best');
            grid on;
            set(gcf,'Color','white');
            
            % Plot R0
            figure;
            %clf;
            hold on;
            % Manual
            plot(R0data.SOC, R0data.R0Discharge, 'LineWidth', 2, 'DisplayName', 'R0 Discharge');
            plot(R0data.SOC, R0data.R0DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R0 Discharge Relaxation');
            plot(R0data.SOC, R0data.R0Charge, 'LineWidth', 2, 'DisplayName', 'R0 Charge');
            plot(R0data.SOC, R0data.R0ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R0 Charge Relaxation');
            % Auto
            plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeR0, 'LineWidth', 2, 'DisplayName', 'R0 Discharge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeR0, 'LineWidth', 2, 'DisplayName', 'R0 Discharge CURVEFIT');
            plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeR0, 'LineWidth', 2, 'DisplayName', 'R0 Charge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeR0, 'LineWidth', 2, 'DisplayName', 'R0 Charge CURVEFIT'); 
            hold off;
            xlabel('SOC (%)');
            xlim([10 90]);
            ylabel('R0 (Ohm)');
            title(['R0 vs SOC - Test',fields{i}], 'Interpreter','none');
            legend('Location','best');
            grid on;
            set(gcf,'Color','white');

            % Plot R1
            figure;
            hold on;
            % Manual
            plot(R1data.SOC, R1data.R1Discharge, 'LineWidth', 2, 'DisplayName', 'R1 Discharge');
            plot(R1data.SOC, R1data.R1DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R1 Discharge Relaxation');
            plot(R1data.SOC, R1data.R1Charge, 'LineWidth', 2, 'DisplayName', 'R1 Charge');
            plot(R1data.SOC, R1data.R1ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R1 Charge Relaxation');
            % Auto
            plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeR1, 'LineWidth', 2, 'DisplayName', 'R1 Discharge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeR1, 'LineWidth', 2, 'DisplayName', 'R1 Discharge CURVEFIT');
            plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeR1, 'LineWidth', 2, 'DisplayName', 'R1 Charge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeR1, 'LineWidth', 2, 'DisplayName', 'R1 Charge CURVEFIT'); 
            hold off;
            xlabel('SOC (%)');
            xlim([10 90]);
            ylabel('R1 (Ohm)');
            title(['R1 vs SOC - Test ', fields{i}], 'Interpreter','none');
            legend('Location','best');
            grid on;
            set(gcf,'Color','white');

            % Plot C1
            figure;
            hold on;
            % Manual
            plot(C1data.SOC, C1data.C1Discharge, 'LineWidth', 2, 'DisplayName', 'C1 Discharge');
            plot(C1data.SOC, C1data.C1DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'C1 Discharge Relaxation');
            plot(C1data.SOC, C1data.C1Charge, 'LineWidth', 2, 'DisplayName', 'C1 Charge');
            plot(C1data.SOC, C1data.C1ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'C1 Charge Relaxation');
            % Auto
            plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeC1, 'LineWidth', 2, 'DisplayName', 'C1 Discharge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeC1, 'LineWidth', 2, 'DisplayName', 'C1 Discharge CURVEFIT');
            plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeC1, 'LineWidth', 2, 'DisplayName', 'C1 Charge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeC1, 'LineWidth', 2, 'DisplayName', 'C1 Charge CURVEFIT'); 
            hold off;
            xlabel('SOC (%)');
            xlim([10 90]);
            ylabel('C1 (Farad)');
            title(['C1 vs SOC - Test ', fields{i}], 'Interpreter','none');
            legend('Location','best');
            grid on;
            set(gcf,'Color','white');

            % Plot R2
            figure;
            hold on;
            % Manual
            plot(R2data.SOC, R2data.R2Discharge, 'LineWidth', 2, 'DisplayName', 'R2 Discharge');
            plot(R2data.SOC, R2data.R2DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R2 Discharge Relaxation');
            plot(R2data.SOC, R2data.R2Charge, 'LineWidth', 2, 'DisplayName', 'R2 Charge');
            plot(R2data.SOC, R2data.R2ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'R2 Charge Relaxation');
            % Auto
            plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeR2, 'LineWidth', 2, 'DisplayName', 'R2 Discharge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeR2, 'LineWidth', 2, 'DisplayName', 'R2 Discharge CURVEFIT');
            plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeR2, 'LineWidth', 2, 'DisplayName', 'R2 Charge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeR2, 'LineWidth', 2, 'DisplayName', 'R2 Charge CURVEFIT'); 
            hold off;
            xlabel('SOC (%)');
            xlim([10 90]);
            ylabel('R2 (Ohm)');
            title(['R2 vs SOC - Test ', fields{i}], 'Interpreter','none');
            legend('Location','best');
            grid on;
            set(gcf,'Color','white');

            % Plot C2
            figure;
            hold on;
            % Manual
            plot(C2data.SOC, C2data.C2Discharge, 'LineWidth', 2, 'DisplayName', 'C2 Discharge');
            plot(C2data.SOC, C2data.C2DischargeRelaxation, 'LineWidth', 2, 'DisplayName', 'C2 Discharge Relaxation');
            plot(C2data.SOC, C2data.C2Charge, 'LineWidth', 2, 'DisplayName', 'C2 Charge');
            plot(C2data.SOC, C2data.C2ChargeRelaxation, 'LineWidth', 2, 'DisplayName', 'C2 Charge Relaxation');
            % Auto
            plot(SOCvector, batteryEcmFM.ModelParameterTables.DischargeC2, 'LineWidth', 2, 'DisplayName', 'C2 Discharge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.DischargeC2, 'LineWidth', 2, 'DisplayName', 'C2 Discharge CURVEFIT');
            plot(SOCvector, batteryEcmFM.ModelParameterTables.ChargeC2, 'LineWidth', 2, 'DisplayName', 'C2 Charge FMINSEARCH');
            plot(SOCvector, batteryEcmCF.ModelParameterTables.ChargeC2, 'LineWidth', 2, 'DisplayName', 'C2 Charge CURVEFIT'); 
            hold off;
            xlabel('SOC (%)');
            xlim([10 90]);
            ylabel('C2 (Farad)');
            title(['C2 vs SOC - Test ', fields{i}], 'Interpreter','none');
            legend('Location','best');
            grid on;
            set(gcf,'Color','white');

            case 14
            %% ESTIMATE HEAT CAPACITY
            % Unite NEWARE and DEWESOFT data
            % Initialize the combined struct (merge of NEWARE and DEWESOFT Data)
            DataCOMBINED = struct();
            DataTCAPACITY = struct();

            % Get the field names of both structs
            fieldsNEWARE = fieldnames(DataNEWARE);
            fieldsDEWESOFT = fieldnames(DataDEWESOFT);

            % Loop over the fields of DataNEWARE to find corresponding fields in DataDEWESOFT
            for i = 1:length(fieldsNEWARE)
                % Check if the current field name ends with 'NEWARE'
                if endsWith(fieldsNEWARE{i}, '_NEWARE_xlsx')
                    % Extract the base name by removing 'NEWARE'
                    baseName = extractBefore(fieldsNEWARE{i}, '_NEWARE_xlsx');
                    % Check if the same base name exists in DataDEWESOFT with 'DEWESOFT'
                    correspondingField = strcat(baseName, '_DEWESOFT_xlsx');
                    if isfield(DataDEWESOFT, correspondingField)
                        % Get the tables from both structs
                        tableNEWARE = DataNEWARE.(fieldsNEWARE{i});
                        tableDEWESOFT = DataDEWESOFT.(correspondingField);

                        % Rename 'Time' to 'Time(s)' for DEWESOFT
                        tableDEWESOFT.Properties.VariableNames{strcmp(tableDEWESOFT.Properties.VariableNames, 'Time')} = 'Time(s)';
                        
                        % Merge the tables using an outer join on 'Time(s)'
                        combinedTable = outerjoin(tableNEWARE, tableDEWESOFT, 'Keys', 'Time(s)', 'MergeKeys', true);

                        % Fill Nan and empty strings with data
                        filledTable = combinedTable;
                        varNames = filledTable.Properties.VariableNames;
                    
                        for k = 1:numel(varNames)
                            colData = filledTable.(varNames{k});
                    
                            % Handle numeric columns
                            if isnumeric(colData)
                                % Forward fill
                                lastVal = NaN;
                                for j = 1:length(colData)
                                    if ~isnan(colData(j))
                                        lastVal = colData(j);
                                    elseif ~isnan(lastVal)
                                        colData(j) = lastVal;
                                    end
                                end
                                % Fill top NaNs with first valid value
                                if isnan(colData(1))
                                    firstValid = find(~isnan(colData), 1, 'first');
                                    if ~isempty(firstValid)
                                        colData(1:firstValid-1) = colData(firstValid);
                                    end
                                end
                                filledTable.(varNames{k}) = colData;
                    
                            % Handle string/cell string columns
                            elseif iscellstr(colData) || isstring(colData)
                                % Convert to cellstr if it's string for uniform handling
                                if isstring(colData)
                                    colData = cellstr(colData);
                                    convertBackToString = true;
                                else
                                    convertBackToString = false;
                                end
                    
                                lastStr = '';
                                for j = 1:length(colData)
                                    if ~strcmp(colData{j}, '')
                                        lastStr = colData{j};
                                    elseif ~strcmp(lastStr, '')
                                        colData{j} = lastStr;
                                    end
                                end
                    
                                % Fill top empty strings with first non-empty
                                if strcmp(colData{1}, '')
                                    firstValid = find(~cellfun(@isempty, colData), 1, 'first');
                                    if ~isempty(firstValid)
                                        colData(1:firstValid-1) = repmat(colData(firstValid), firstValid-1, 1);
                                    end
                                end
                    
                                % Convert back if necessary
                                if convertBackToString
                                    colData = string(colData);
                                end
                    
                                filledTable.(varNames{k}) = colData;
                            end
                        end

                        % Keep only rest phase after pilot discharge by considering data from the 3rd 'Rest' phase only
                        is_rest = strcmp(filledTable.('Step Type'), 'Rest'); % Logical index of 'Rest' rows
                        transitions = diff([0; is_rest; 0]);  % Find the start indices of all 'Rest' blocks
                        starts = find(transitions == 1);      % start of each block
                        
                        start_idx = starts(3); % Get the start index of the 3rd block
                        trimmedTable = filledTable(start_idx:end, :); % Keep the table from that point onward
                        trimmedTable.('Time(s)') = trimmedTable.('Time(s)')-trimmedTable.('Time(s)')(1); % Adjust time
                        
                        is_rest = strcmp(trimmedTable.('Step Type'), 'Rest'); % Logical index of 'Rest' rows
                        transitions = diff([0; is_rest; 0]);  % Find the start indices of all 'Rest' blocks
                        ends   = find(transitions == -1) - 1; % end of each block

                        cut_idx = ends(1);  % Get the end index of the 1st block  
                        trimmedTable2 = trimmedTable(1:cut_idx, :); % Keep only rows up to that point

                        % Save the combined table in the new struct
                        DataCOMBINED.(baseName) = trimmedTable2;

                        % Save the combined table in the new struct (Only useful data)
                        DataTCAPACITY.(baseName) = DataCOMBINED.(baseName)(:, {'Time(s)','Temperature (C)','Heat Flux Density (W/m^2)'});
                        DataTCAPACITY.(baseName).Properties.VariableNames{strcmp(DataTCAPACITY.(baseName).Properties.VariableNames, 'Time(s)')} = 'Time (s)';
                    end
                end
            end

            fields=fieldnames(DataTCAPACITY);
            Cth=struct();
            h=struct;

            for i = 1:length(fields) % Iterating all the tests
                Data=DataTCAPACITY.(fields{i});
                
                % Thermal capacity calculation
                t=Data.('Time (s)'); % t: time vector in seconds
                t=t(2400:end-1800); % delete first 40 minutes and last 30 minutes
                T=Data.('Temperature (C)'); % T: temperature vector in Celsius
                T=T(2400:end-1800); % delete first 40 minutes and last 30 minutes
                q=Data.('Heat Flux Density (W/m^2)'); % q: heat flux density vector in W/m^2
                q=q(2400:end-1800); % delete first 40 minutes and last 30 minutes
                
                % Define Cell Surface Area (m)
                A = pi*21*70*10^(-6)+2*pi*(21/2)^2*10^(-6);

                % Define Cell Mass (kg)
                m = 0.07;
                
                % Calculate Total Heat Power (W)
                Q_out = q * A;
                
                % Smooth the Temperature Data (disabled as smoothed already previously)
                % 'sgolay' preserves features better than a simple moving average (adjust the window size based on your sampling rate and noise)
                % window_size = 25; 
                % T_smooth = smoothdata(T, 'sgolay', window_size);
                T_smooth=T;
                
                % Calculate the Temperature Derivative (dT/dt)
                % Using 'gradient' is preferable to 'diff' because it returns a vector
                dTdt = gradient(T_smooth, t);
                
                % Define the X-variable (negative dT/dt, since the cell is cooling)
                x_var = -dTdt;
                
                % Perform Zero-Intercept Linear Regression
                % use fitlm and explicitly set 'Intercept' to false to force the line through (0,0)
                mdl = fitlm(x_var, Q_out, 'Intercept', false);
                
                % Extract the slope, which is the Thermal Capacity (C_th)
                C_th = mdl.Coefficients.Estimate;
                c = C_th/m;
                Rsquared = mdl.Rsquared.Ordinary;
                
                % Display the result in the command window
                fprintf('Test %s \n', fields{i});
                fprintf('Estimated Thermal Capacity (C_th) = %.2f J/K\n', C_th);
                fprintf('Estimated Specific Heat Capacity (c) = %.2f J/(kg K)\n', c);
                fprintf('R squared = %.2f \n', Rsquared);
                
                % Visualize the Regression
                figure;
                scatter(x_var, Q_out, 'b.', 'DisplayName', 'Experimental Data');
                hold on;
                plot(x_var, C_th * x_var, 'r-', 'LineWidth', 2, 'DisplayName', sprintf('Fit: c = %.2f J/(kg K)', c));
                xlabel('-dT/dt (K/s)');
                ylabel('{Q}_{out} (W)', 'Interpreter', 'tex');
                title(['Thermal Capacity Estimation - Test ', fields{i}], 'Interpreter','none');
                legend('Location', 'best');
                grid on;
                xlim([0 inf]);
                ylim([0 inf]);
                set(gcf,'Color','white');
                
                % Result validation (Biot number)
                % Define ambient temperature (C)
                T_amb = T(end);
                
                % Define characteristic length (m)
                L_c = 21/2*70/2/(21/2+70)*10^(-3);
                
                % Estimated through-plane thermal conductivity (W/m*K)
                k = 1.0;             
                
                % Calculate the Convective Heat Transfer Coefficient (h)
                % We only want to calculate h when there is a meaningful temperature difference (e.g., delta T > 1 degree) to avoid dividing by noise near zero.
                delta_T = T - T_amb;
                
                % Create a logical mask to find where the temperature difference is significant
                valid_indices = delta_T > 1.0; 
                
                if sum(valid_indices) == 0
                    warning('Temperature difference is too small to accurately calculate h.');
                    h_avg = NaN;
                else
                    % Calculate h (W/m^2*K) only for the valid data points
                    h_instantaneous = q(valid_indices) ./ delta_T(valid_indices);
                    
                    % Take the mean to get a single, stable value for the environment
                    h_avg = mean(h_instantaneous);
                end
                
                % Calculate the Biot Number
                Bi = (h_avg * L_c) / k;

                % Calculate R_th
                R_th = 1/(h_avg * A);
                
                % Display Results
                fprintf('Estimated Average h: %.2f W/(m^2 K)\n', h_avg);
                fprintf('Calculated Biot Number (Bi): %.4f\n', Bi);
                fprintf('Calculated Thermal Resistance (R_th): %.2f K/W \n', R_th);
                
                if Bi < 0.1
                    disp('Bi < 0.1: The lumped capacitance assumption is valid.');
                else
                    disp('Bi >= 0.1: The lumped capacitance assumption is NOT strictly valid.');
                end

                % Final save
                Cth.(fields{i})=C_th;
                h.(fields{i})=h_avg;
            end

        case 15
            %% EXPORT DATA FOR SIMULINK
            % Unite NEWARE and DEWESOFT data
            % Initialize the combined struct (merge of NEWARE and DEWESOFT Data)
            DataCOMBINED = struct();
            DataHPPC = struct();

            % Get the field names of both structs
            fieldsNEWARE = fieldnames(DataNEWARE);
            fieldsDEWESOFT = fieldnames(DataDEWESOFT);

            % Loop over the fields of DataNEWARE to find corresponding fields in DataDEWESOFT
            for i = 1:length(fieldsNEWARE)
                % Check if the current field name ends with 'NEWARE'
                if endsWith(fieldsNEWARE{i}, '_NEWARE_xlsx')
                    % Extract the base name by removing 'NEWARE'
                    baseName = extractBefore(fieldsNEWARE{i}, '_NEWARE_xlsx');
                    % Check if the same base name exists in DataDEWESOFT with 'DEWESOFT'
                    correspondingField = strcat(baseName, '_DEWESOFT_xlsx');
                    if isfield(DataDEWESOFT, correspondingField)
                        % Get the tables from both structs
                        tableNEWARE = DataNEWARE.(fieldsNEWARE{i});
                        tableDEWESOFT = DataDEWESOFT.(correspondingField);

                        % Rename 'Time' to 'Time(s)' for DEWESOFT
                        tableDEWESOFT.Properties.VariableNames{strcmp(tableDEWESOFT.Properties.VariableNames, 'Time')} = 'Time(s)';
                        
                        % Merge the tables using an outer join on 'Time(s)'
                        combinedTable = outerjoin(tableNEWARE, tableDEWESOFT, 'Keys', 'Time(s)', 'MergeKeys', true);

                        % Fill Nan and empty strings with data
                        filledTable = combinedTable;
                        varNames = filledTable.Properties.VariableNames;
                    
                        for k = 1:numel(varNames)
                            colData = filledTable.(varNames{k});
                    
                            % Handle numeric columns
                            if isnumeric(colData)
                                % Forward fill
                                lastVal = NaN;
                                for j = 1:length(colData)
                                    if ~isnan(colData(j))
                                        lastVal = colData(j);
                                    elseif ~isnan(lastVal)
                                        colData(j) = lastVal;
                                    end
                                end
                                % Fill top NaNs with first valid value
                                if isnan(colData(1))
                                    firstValid = find(~isnan(colData), 1, 'first');
                                    if ~isempty(firstValid)
                                        colData(1:firstValid-1) = colData(firstValid);
                                    end
                                end
                                filledTable.(varNames{k}) = colData;
                    
                            % Handle string/cell string columns
                            elseif iscellstr(colData) || isstring(colData)
                                % Convert to cellstr if it's string for uniform handling
                                if isstring(colData)
                                    colData = cellstr(colData);
                                    convertBackToString = true;
                                else
                                    convertBackToString = false;
                                end
                    
                                lastStr = '';
                                for j = 1:length(colData)
                                    if ~strcmp(colData{j}, '')
                                        lastStr = colData{j};
                                    elseif ~strcmp(lastStr, '')
                                        colData{j} = lastStr;
                                    end
                                end
                    
                                % Fill top empty strings with first non-empty
                                if strcmp(colData{1}, '')
                                    firstValid = find(~cellfun(@isempty, colData), 1, 'first');
                                    if ~isempty(firstValid)
                                        colData(1:firstValid-1) = repmat(colData(firstValid), firstValid-1, 1);
                                    end
                                end
                    
                                % Convert back if necessary
                                if convertBackToString
                                    colData = string(colData);
                                end
                    
                                filledTable.(varNames{k}) = colData;
                            end
                        end

                        % Keep only HPPC test (delete inititial charge/discharge) by starting to consider data from the 4th 'Rest' phase
                        is_rest = strcmp(filledTable.('Step Type'), 'Rest'); % Logical index of 'Rest' rows
                        transitions = diff([0; is_rest; 0]);  % Find the start indices of all 'Rest' blocks
                        starts = find(transitions == 1);      % start of each block
                        ends   = find(transitions == -1) - 1; % end of each block (not used here)
                        if length(starts) < 4 % Check if there are at least 4 'Rest' blocks
                            disp('Less than 4 ''Rest'' blocks found.');
                            trimmedTable = table();  % return empty table
                        else
                            start_idx = starts(4); % Get the start index of the 4th block
                            trimmedTable = filledTable(start_idx:end, :); % Keep the table from that point onward
                            trimmedTable.('Time(s)') = trimmedTable.('Time(s)')-trimmedTable.('Time(s)')(1); % Adjust time
                        end

                        % Keep only HPPC test (delete final discharge) by finishing to consider up to the second to last 'Rest' phase
                        is_rest = strcmp(trimmedTable.('Step Type'), 'Rest'); % Logical index of 'Rest' rows
                        transitions = diff([0; is_rest; 0]); % Identify block starts and ends
                        starts = find(transitions == 1);
                        ends   = find(transitions == -1) - 1;
                        if length(starts) < 2 % Ensure there are at least 2 blocks
                            disp('Less than two ''Rest'' blocks found.');
                            trimmedTable2 = trimmedTable;  % return original table
                        else
                            cut_idx = ends(end-1);  % Get the end index of the second-to-last block  
                            trimmedTable2 = trimmedTable(1:cut_idx, :); % Keep only rows up to that point
                        end

                        % Save the combined table in the new struct
                        DataCOMBINED.(baseName) = trimmedTable2;

                        % Save the combined table in the new struct (Only useful data)
                        DataHPPC.(baseName) = DataCOMBINED.(baseName)(:, {'Step Type','Time(s)', 'Current(A)', 'Voltage(V)', 'Temperature (C)','SOC', 'Heat Flux Density (W/m^2)'});
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Time(s)')} = 'Time (s)';
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Current(A)')} = 'Current (A)';
                        DataHPPC.(baseName).Properties.VariableNames{strcmp(DataHPPC.(baseName).Properties.VariableNames, 'Voltage(V)')} = 'Voltage (V)';
                    end
                end
            end
            % Get available table names
            tableNames = fieldnames(DataHPPC);
            
            % Let the user select a NEWARE table
            [tableIndex, okTable] = listdlg('ListString', tableNames, ...
                                            'SelectionMode', 'single', ...
                                            'PromptString', 'Select a table:', ...
                                            'ListSize', [200 150]);
            
            % Check if user made a selection
            if okTable == 0
                disp('No table selected. Exiting.');
                return;
            end

            % Get the selected tables
            selectedTableName = tableNames{tableIndex};
            selectedTable = DataHPPC.(selectedTableName);

            % Export Current
            Current_Profile = timeseries(selectedTable.('Current (A)'), selectedTable.('Time (s)'));
            save('Current_Profile.mat', 'Current_Profile');

            % Export Voltage
            Voltage_Profile = timeseries(selectedTable.('Voltage (V)'), selectedTable.('Time (s)'));
            save('Voltage_Profile.mat', 'Voltage_Profile');

            % Export Temperature
            Temperature_Profile = timeseries(selectedTable.('Temperature (C)'), selectedTable.('Time (s)'));
            save('Temperature_Profile.mat', 'Temperature_Profile');

            % Export Heat Flux Density
            Heat_Flux_Density_Profile = timeseries(selectedTable.('Heat Flux Density (W/m^2)'), selectedTable.('Time (s)'));
            save('Heat_Flux_Density_Profile.mat', 'Heat_Flux_Density_Profile');

            disp('Success: Data exported to .mat files');

        case 16
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

        case 17
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

        case 18
            %% CLOSE
            exit=true;
        otherwise
            exit=true;
    end

end
