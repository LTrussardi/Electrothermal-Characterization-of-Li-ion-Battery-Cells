%close all;
%clc;
%clear;

% CELL TESTING DATA ANALYSIS MATLAB CODE
% Created by Luca Trussardi in March 2025

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
             "9) Plot Discharge Voltage VS Capacity",...
             "10) Plot Voltage/Current Charging Phase",...
             "11) Plot Voltage/Currrent Discharging Phase",...
             "12) Plot Temperature/Heat Flux Discharging Phase",...
             "13) Estimate Entropic Coefficient",...
             "14) Export Data for SIMULINK",...
             "S) Save Data to .mat files",...
             "L) Load Data from .mat files",...
             "C) Close"];
    choice=menu(message,options); % Creates menu window that allows to select what to do

    switch choice
        case 1
            %% IMPORT NEWARE
            DataNEWARE=ImportFiles('record','DataNEWARE');

        case 2
            %% IMPORT DEWESOFT
            DataDEWESOFT=ImportFiles('Data1','DataDEWESOFT','Data1-1');
            DataDEWESOFTfields = fieldnames(DataDEWESOFT);
            for i = 1:length(DataDEWESOFTfields)
                DataDEWESOFT.(DataDEWESOFTfields{i}) = DataDEWESOFT.(DataDEWESOFTfields{i})(1:10:end, :);
                if any(strcmp(DataDEWESOFT.(DataDEWESOFTfields{i}).Properties.VariableNames, "Thermocouple K/AVE"))
                    DataDEWESOFT.(DataDEWESOFTfields{i})(:, "Thermocouple K/AVE") = []; % Remove the column
                end
                if any(strcmp(DataDEWESOFT.(DataDEWESOFTfields{i}).Properties.VariableNames, "Q(t) - Sample heat"))
                    DataDEWESOFT.(DataDEWESOFTfields{i})(:, "Q(t) - Sample heat") = []; % Remove the column
                end
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
            defaultSelectionN = [5,6,7,8,10];
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
            TimeVector=(0:0.1:(height(selectedTableN)-1)/10)';
            
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
            defaultSelectionD = [3,4,5,6];
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
            
            % Create time vector
            maxheight = 0;
            for i = 1:length(selectedTables)
                T = DataNEWARE.(selectedTables{i});
                if height(T) > maxheight
                    maxheight = height(T);
                end   
            end
            TimeVector = (0:0.1:(maxheight-1)/10)';
            
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
            defaultSelectionN = [5,6,7,8,10];
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
            defaultSelectionD = [3,4,5,6];
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
            TimeVectorN=(0:0.1:(height(selectedTableN)-1)/10)';
            TimeVectorD=(0:0.1:(height(selectedTableD)-1)/10)';
            if height(selectedTableN)>height(selectedTableD)
                TimeVector=TimeVectorN;
            else
                TimeVector=TimeVectorD;
            end
            
            plotchoice=menu("How do you want to plot?","Single figure with subplots","Separate figures","Single plot with multiple axes");
            
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
                case 3
                    % Join data
                    selectedTableN.RowNum = (1:height(selectedTableN))'; % Create row index
                    selectedTableD.RowNum = (1:height(selectedTableD))';

                    selectedColsN{end+1}='RowNum';
                    selectedColsD{end+1}='RowNum';
                    
                    selectedData = outerjoin(selectedTableN(:,selectedColsN), selectedTableD(:,selectedColsD), 'Keys', 'RowNum', 'MergeKeys', true);
                    selectedData.RowNum = []; % Remove the helper column
                    %selectedData=[selectedTableN(:,selectedColsN),selectedTableD(:,selectedColsD)];
                    % Create a figure with multiple axes
                    figure;
                    ax1 = axes; % Create the first axes object
                    numPlots = length(selectedColsN)+length(selectedColsD)-2;
                    colors = lines(numPlots); % Generate distinct colors

                    % Create an array to hold all axes
                    allAxes = ax1;

                    % Loop through selected columns and create multiple y-axes
                    columnNamesTot = selectedData.Properties.VariableNames;
                    hold on;
                    for i = 1:numPlots
                        if i == 1
                            % First plot on main axes (ax1)
                            plot(ax1, TimeVector, selectedData{:, columnNamesTot{i}}, 'Color', colors(i, :), 'LineWidth', 1.5);
                            ax1.YColor = colors(i, :);
                            ylabel(ax1, columnNamesTot{i});
                        else
                            % Create new axes on top of existing one
                            axNew = axes('Position', ax1.Position, 'Color', 'none', 'YAxisLocation', 'right');
                            hold on;
                            plot(axNew, TimeVector, selectedData{:, columnNamesTot{i}}, 'Color', colors(i, :), 'LineWidth', 1.5);
                            axNew.YColor = colors(i, :);
                            ylabel(axNew, columnNamesTot{i});
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

        case 9
            %% PLOT DISCHARGE VOLTAGE VS CAPACITY
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

            % Extract data during the discharge phase
            DischargeDataNEWARE=DataNEWARE;
            for i = 1:length(selectedTables)
                T=DischargeDataNEWARE.(selectedTables{i});
                dischargeT = T(strcmp(T.("Step Type"), 'CC DChg'), :);
                DischargeDataNEWARE.(selectedTables{i})=dischargeT;
            end

            % Create figure
            figure('Name', 'Voltage VS Capacity', 'NumberTitle', 'off');
            hold on;
            numPlots = length(selectedTables);
            for i = 1:numPlots
                plot(DischargeDataNEWARE.(selectedTables{i}).("Capacity(Ah)"), DischargeDataNEWARE.(selectedTables{i}).("Voltage(V)"), 'LineWidth', 1.5, 'DisplayName', selectedTables{i});
            end
            hold off;
            xlabel('Capacity(Ah)');
            ylabel('Voltage(V)');
            title('Plot of Voltage(V) vs Capacity(Ah)');
            leg=legend;
            set(leg,'Interpreter', 'none');
            grid on;

        case 10
            %% PLOT CHARGING PHASE
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

            % Extract data during the charge phase
            ChargeDataNEWARE=selectedTableN(strcmp(selectedTableN.("Step Type"), 'CCCV Chg'), :);

            % Create the time vector
            TimeVector=(0:0.1:(height(ChargeDataNEWARE)-1)/10)';
                        
            % Create figure
            figure('Name', 'Charging Voltage(V) and Current(A) VS Time(s)', 'NumberTitle', 'off');
            yyaxis left;
            plot(TimeVector,ChargeDataNEWARE.("Voltage(V)"), 'LineWidth', 1.5, 'DisplayName', 'Voltage');
            ylabel('Voltage(V)');
            yyaxis right;
            plot(TimeVector,ChargeDataNEWARE.("Current(A)"), 'LineWidth', 1.5, 'DisplayName', 'Current');
            xlabel('Time(s)');
            title('Plot of Voltage(V) and Current(A) VS Time(s) during charging');
            legend;
            grid on;

            case 11
            %% PLOT DISCHARGING PHASE
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

            % Extract data during the discharge phase
            ChargeDataNEWARE=selectedTableN(strcmp(selectedTableN.("Step Type"), 'CC DChg'), :);

            % Create the time vector
            TimeVector=(0:0.1:(height(ChargeDataNEWARE)-1)/10)';
                        
            % Create figure
            figure('Name', 'Charging Voltage(V) and Current(A) VS Time(s)', 'NumberTitle', 'off');
            yyaxis left;
            plot(TimeVector,ChargeDataNEWARE.("Voltage(V)"), 'LineWidth', 1.5, 'DisplayName', 'Voltage');
            ylabel('Voltage(V)');
            yyaxis right;
            plot(TimeVector,ChargeDataNEWARE.("Current(A)"), 'LineWidth', 1.5, 'DisplayName', 'Current');
            xlabel('Time(s)');
            title('Plot of Voltage(V) and Current(A) VS Time(s) during discharging');
            legend;
            grid on;

        case 12
            %% PLOT TEMPERATURE AND HEAT FLUX FROM DISCHARGING PHASE
            DataNEWARE2=DataNEWARE;
            DataDEWESOFT2=DataDEWESOFT;

            % Fix Time, convert repeated 'hh:mm:ss' timestamps to seconds with 0.1s increments
            tableNames = fieldnames(DataNEWARE2); % Get the available table names
                        
            for i = 1:length(tableNames) % Correct the time vector
                timeStrArray=DataNEWARE2.(tableNames{i}).("Total Time");
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

               DataNEWARE2.(tableNames{i}).("Time(s)") = totalSeconds;
            end

            % Calculate Heat Flux from Voltage and automatically add it to data
            Sensitivity=17.39*10^(-6); % V*m^2/W
            TestArea=12.96*10^(-4); % m^2
            LateralSurfaceArea=2*pi*21*70*10^(-6); % m^2

            tableNames = fieldnames(DataDEWESOFT2); % Get the available table names
            
            % For each of the tests, calculates the heat flux quantities and adds them to the data
            for i = 1:length(tableNames)          
                % Heat fluxes culculation with temperature corrected sensitivity value
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)")=DataDEWESOFT2.(tableNames{i}).("Thermal flux FHF05")./(Sensitivity.*(1+0.002.*(DataDEWESOFT2.(tableNames{i}).("Temperature FHF05")-20))); % W/m^2
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Test Area (W)")=DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)").*TestArea; % W on test area of the sensor
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Lateral Surface Area (W)")=DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)").*LateralSurfaceArea; % W on the lateral surface area

                % Rename 'Thermal flux FHF05' to 'Thermal Flux Sensor Voltage (V)'
                DataDEWESOFT2.(tableNames{i}).Properties.VariableNames{'Thermal flux FHF05'} = 'Thermal Flux Sensor Voltage (V)';

                % Rename 'Temperature FHF05' to 'Temperature (C)'
                DataDEWESOFT2.(tableNames{i}).Properties.VariableNames{'Temperature FHF05'} = 'Temperature (C)';
            end

            % Unite NEWARE and DEWESOFT data
            % Initialize the combined struct (merge of NEWARE and DEWESOFT Data)
            DataCOMBINED = struct();

            % Get the field names of both structs
            fieldsNEWARE = fieldnames(DataNEWARE2);
            fieldsDEWESOFT = fieldnames(DataDEWESOFT2);

            % Loop over the fields of DataNEWARE2 to find corresponding fields in DataDEWESOFT2
            for i = 1:length(fieldsNEWARE)
                % Check if the current field name ends with 'NEWARE'
                if endsWith(fieldsNEWARE{i}, '_NEWARE_xlsx')
                    % Extract the base name by removing 'NEWARE'
                    baseName = extractBefore(fieldsNEWARE{i}, '_NEWARE_xlsx');
                    % Check if the same base name exists in DataDEWESOFT2 with 'DEWESOFT'
                    correspondingField = strcat(baseName, '_DEWESOFT_xlsx');
                    if isfield(DataDEWESOFT2, correspondingField)
                        % Get the tables from both structs
                        tableNEWARE = DataNEWARE2.(fieldsNEWARE{i});
                        tableDEWESOFT = DataDEWESOFT2.(correspondingField);

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

                        % Save the combined table in the new struct
                        DataCOMBINED.(baseName) = filledTable;

                    end
                end
            end


            % Get the available table names
            tableNames = fieldnames(DataCOMBINED);
            
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

            % Extract data during the discharge phase
            DischargeDataCOMBINED=DataCOMBINED;
            for i = 1:length(selectedTables)
                T=DischargeDataCOMBINED.(selectedTables{i});
                rest_idx = strcmp(T.('Step Type'), 'Rest'); % Identify where 'Step Type' is 'Rest'
                transitions = diff([0; rest_idx; 0]); % Find start and end indices of all continuous blocks
                starts = find(transitions == 1);
                ends   = find(transitions == -1) - 1;
                T.('Time(s)')=T.('Time(s)')-T.('Time(s)')(ends(2));
                dischargeT = T(ends(2):end, :);
                DischargeDataCOMBINED.(selectedTables{i})=dischargeT;
            end

            % Create figures
            figure('Name', 'Temperature VS Time', 'NumberTitle', 'off');
            hold on;
            numPlots = length(selectedTables);
            for i = 1:numPlots
                plot(DischargeDataCOMBINED.(selectedTables{i}).("Time(s)"), DischargeDataCOMBINED.(selectedTables{i}).("Temperature (C)"), 'LineWidth', 1.5, 'DisplayName', selectedTables{i});
            end
            hold off;
            xlabel('Time (s)');
            ylabel('Temperature (C)');
            title('Plot of Temperature(C) vs Time(s)');
            leg=legend;
            set(leg,'Interpreter', 'none');
            grid on;

            figure('Name', 'Heat Flux Density VS Time', 'NumberTitle', 'off');
            hold on;
            numPlots = length(selectedTables);
            for i = 1:numPlots
                plot(DischargeDataCOMBINED.(selectedTables{i}).("Time(s)"), DischargeDataCOMBINED.(selectedTables{i}).("Heat Flux Density (W/m^2)"), 'LineWidth', 1.5, 'DisplayName', selectedTables{i});
            end
            hold off;
            xlabel('Time (s)');
            ylabel('Heat Flux Density (W/m^2)');
            title('Plot of Heat Flux Density(W/m^2) vs Time(s)');
            leg=legend;
            set(leg,'Interpreter', 'none');
            grid on;

        case 13
            %% ESTIMATE ENTROPIC COEFFICIENT
            DataNEWARE2=DataNEWARE;
            DataDEWESOFT2=DataDEWESOFT;

            % Fix Time, convert repeated 'hh:mm:ss' timestamps to seconds with 0.1s increments
            tableNames = fieldnames(DataNEWARE2); % Get the available table names
                        
            for i = 1:length(tableNames) % Correct the time vector
                timeStrArray=DataNEWARE2.(tableNames{i}).("Total Time");
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

               DataNEWARE2.(tableNames{i}).("Time(s)") = totalSeconds;
            end

            % Calculate SOC
            for i = 1:length(tableNames)
                if contains(DataNEWARE2.(tableNames{i}).Properties.VariableNames,"SOC")
                else
                    T=DataNEWARE2.(tableNames{i});
                    % Finding the real Capacity curve by integration (the one provided is reset at each phase)
                    Capacity=cumtrapz(DataNEWARE2.(tableNames{i}).("Time(s)"),DataNEWARE2.(tableNames{i}).("Current(A)"))/3600;
                    T.('Real Capacity by Integration(Ah)')=Capacity;
                    
                    % Finding the Capacity corresponding to 0 and 100% SOC
                    rest_idx = strcmp(T.('Step Type'), 'Rest'); % Identify where 'Step Type' is 'Rest'
                    transitions = diff([0; rest_idx; 0]); % Find start and end indices of all continuous blocks
                    starts = find(transitions == 1);
                    ends   = find(transitions == -1) - 1;
                    
                    Capacity0 = T.('Real Capacity by Integration(Ah)')(starts(3)); % Retrieve the corresponding 'Real Capacity by Integration (Ah)'
                    Capacity100 = T.('Real Capacity by Integration(Ah)')(ends(2)); % Retrieve the corresponding 'Real Capacity by Integration (Ah)'
                    
                    % Finding maximum capacity during the pilot discharge phase (used as reference)
                    maxCapacity=Capacity100-Capacity0;
                    % Finding SOC
                    T.('SOC')=(T.('Real Capacity by Integration(Ah)')-Capacity0)./maxCapacity.*100;
                    
                    % Delete useless coloumns and rename some of them for more clarity
                    T(:, {'DataPoint', 'Time','Total Time','Capacity(Ah)','Energy(Wh)','Date','Power(W)'}) = [];
                    T.Properties.VariableNames{strcmp(T.Properties.VariableNames, 'Real Capacity by Integration(Ah)')} = 'Capacity(Ah)';

                    DataNEWARE2.(tableNames{i})=T;
                    DataNEWARE2.(tableNames{i}).('Discharge capacity (Ah)')(:)=maxCapacity; % Save value
                end
            end

            % Calculate Heat Flux from Voltage and automatically add it to data
            Sensitivity=17.39*10^(-6); % V*m^2/W
            TestArea=12.96*10^(-4); % m^2
            LateralSurfaceArea=2*pi*21*70*10^(-6); % m^2

            tableNames = fieldnames(DataDEWESOFT2); % Get the available table names
            
            % For each of the tests, calculates the heat flux quantities and adds them to the data
            for i = 1:length(tableNames)          
                % Heat fluxes culculation with temperature corrected sensitivity value
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)")=DataDEWESOFT2.(tableNames{i}).("Thermal flux FHF05")./(Sensitivity.*(1+0.002.*(DataDEWESOFT2.(tableNames{i}).("Temperature FHF05")-20))); % W/m^2
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Test Area (W)")=DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)").*TestArea; % W on test area of the sensor
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Lateral Surface Area (W)")=DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)").*LateralSurfaceArea; % W on the lateral surface area

                % Rename 'Thermal flux FHF05' to 'Thermal Flux Sensor Voltage (V)'
                DataDEWESOFT2.(tableNames{i}).Properties.VariableNames{'Thermal flux FHF05'} = 'Thermal Flux Sensor Voltage (V)';

                % Rename 'Temperature FHF05' to 'Temperature (C)'
                DataDEWESOFT2.(tableNames{i}).Properties.VariableNames{'Temperature FHF05'} = 'Temperature (C)';
            end

            % Unite NEWARE and DEWESOFT data
            % Initialize the combined struct (merge of NEWARE and DEWESOFT Data)
            DataCOMBINED = struct();

            % Get the field names of both structs
            fieldsNEWARE = fieldnames(DataNEWARE2);
            fieldsDEWESOFT = fieldnames(DataDEWESOFT2);

            % Loop over the fields of DataNEWARE2 to find corresponding fields in DataDEWESOFT2
            for i = 1:length(fieldsNEWARE)
                % Check if the current field name ends with 'NEWARE'
                if endsWith(fieldsNEWARE{i}, '_NEWARE_xlsx')
                    % Extract the base name by removing 'NEWARE'
                    baseName = extractBefore(fieldsNEWARE{i}, '_NEWARE_xlsx');
                    % Check if the same base name exists in DataDEWESOFT2 with 'DEWESOFT'
                    correspondingField = strcat(baseName, '_DEWESOFT_xlsx');
                    if isfield(DataDEWESOFT2, correspondingField)
                        % Get the tables from both structs
                        tableNEWARE = DataNEWARE2.(fieldsNEWARE{i});
                        tableDEWESOFT = DataDEWESOFT2.(correspondingField);

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

                        % Save the combined table in the new struct
                        DataCOMBINED.(baseName) = filledTable;

                    end
                end
            end

            % Get the available table names
            tableNames = fieldnames(DataCOMBINED);
            
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

            ehc=struct;
            ehc.('SOC')=(0:0.01:1)'*100;

            for i = 1:length(selectedTables)

                T=DataCOMBINED.(selectedTables{i});

                % Extract data during the discharge phase
                DischargeData=T(strcmp(T.("Step Type"), 'CC DChg'), :);
            
                % Cell properties
                A = pi*21*70*10^(-6)+2*pi*(21/2)^2*10^(-6);  % Total surface area (m^2)
                C_th = 63.266; % Total thermal capacity (J/K)
    
                % Inputs
                OCV=load('OCV.mat');
                OCVdata=OCV.OCV.('HPPC_1C_21_05_25');
                t_exp=DischargeData.('Time(s)')-DischargeData.('Time(s)')(1); % Time array (seconds)
                T_s_exp=DischargeData.('Temperature (C)'); % Measured surface temperature (C)
                I_exp=DischargeData.('Current(A)'); % Current (A)
                V_exp=DischargeData.('Voltage(V)'); % Measured voltage (V)
                Q_out_exp=DischargeData.('Heat Flux Density (W/m^2)')*A; % Sensor heat flux (W) (q * A_sensor)
                SOC_array=DischargeData.('SOC'); % State of Charge
                InstantOCV=interp1(OCVdata.SOC, OCVdata.OCV, SOC_array, 'linear', 'extrap'); % OCV interpolation (V)
                
                % Calculate the Temperature Gradient (dT/dt)
                dT_dt = gradient(T_s_exp, t_exp);

                % Cap outliers using a moving median (avoids spikes)
                window_size = 50; % Adjust this based on noise frequency
                dT_dt = filloutliers(dT_dt, 'clip', 'movmedian', window_size);
                
                % Calculate Total Heat Generation
                Q_gen_tot = C_th .* dT_dt + Q_out_exp;
                
                % Calculate Irreversible Joule Heat
                Q_joule = I_exp .* (V_exp - InstantOCV);
                
                % Isolate the Reversible (Entropic) Heat
                Q_rev = Q_gen_tot - Q_joule;
                
                % Calculate the Entropic Coefficient (dU/dT)
                % Equation: Q_rev = I * T * (dU/dT) Rearranged: dU/dT = Q_rev / (I * T)
                T_kelvin = T_s_exp + 273.15;
                dU_dT = Q_rev ./ (I_exp .* T_kelvin);
                
                % Smooth before plotting
                smoothingfactor = round(t_exp(end));
                Q_gen_tot = smoothdata(Q_gen_tot, 'movmean', smoothingfactor);
                Q_joule = smoothdata(Q_joule, 'movmean', smoothingfactor);
                Q_rev = smoothdata(Q_rev, 'movmean', smoothingfactor);
                dU_dT = smoothdata(dU_dT, 'sgolay', smoothingfactor);

                % Normalize to standard SOC vector
                SOC_standard = (0:0.01:1)*100;
                [SOC_unique, unique_idx] = unique(SOC_array, 'stable');
                dU_dT_unique = dU_dT(unique_idx);
                dU_dT_standard = interp1(SOC_unique, dU_dT_unique, SOC_standard, 'linear', 'extrap');
                
                % Plotting
                figure;
                plot(SOC_array, Q_gen_tot, 'k-', 'LineWidth', 2, 'DisplayName', 'Total Heat (Joule + Entropic)');
                hold on;
                plot(SOC_array, Q_joule, 'r--', 'LineWidth', 1.5, 'DisplayName', 'Joule Heat');
                plot(SOC_array, Q_rev, 'b-', 'LineWidth', 1.5, 'DisplayName', 'Entropic Heat');
                %set(gca, 'XDir', 'reverse'); % Plot SOC from 100% down to 0%
                xlabel('SOC (%)');
                ylabel('Heat Generation (W)');
                title(["Heat Source Separation - Test ",selectedTables{i}], 'Interpreter','none');
                legend('Location', 'best');
                grid on;
                %ylim([-2 2]);
                
                figure;
                plot(SOC_array, dU_dT * 1000, 'm-', 'LineWidth', 2); % Multiplied by 1000 for mV/K
                %set(gca, 'XDir', 'reverse');
                xlabel('SOC (%)');
                ylabel('EHC (mV/K)');
                title(['Entropic Coefficient - Test ',selectedTables{i}], 'Interpreter','none');
                grid on;
                ylim([-0.25 0.25]);

                ehc.(selectedTables{i})=dU_dT_standard';
            end

            % Comparison plot
            figure('Name', 'EHC VS SOC', 'NumberTitle', 'off');
            hold on;
            numPlots = length(selectedTables);
            for i = 1:numPlots
                plot(SOC_standard, ehc.(selectedTables{i})*1000, 'LineWidth', 1.5, 'DisplayName', selectedTables{i});
            end
            hold off;
            xlabel('SOC (%)');
            ylabel('EHC (mV/K)');
            title('Plot of EHC (mV/K) vs SOC (%)');
            leg=legend;
            set(leg,'Interpreter', 'none');
            grid on;
            ylim([-0.4 0.4]);

        case 14
            %% EXPORT DATA FOR SIMULINK
            DataNEWARE2=DataNEWARE;
            DataDEWESOFT2=DataDEWESOFT;

            % Fix Time, convert repeated 'hh:mm:ss' timestamps to seconds with 0.1s increments
            tableNames = fieldnames(DataNEWARE2); % Get the available table names
                        
            for i = 1:length(tableNames) % Correct the time vector
                timeStrArray=DataNEWARE2.(tableNames{i}).("Total Time");
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

               DataNEWARE2.(tableNames{i}).("Time(s)") = totalSeconds;
            end

            % Calculate SOC
            for i = 1:length(tableNames)
                if contains(DataNEWARE2.(tableNames{i}).Properties.VariableNames,"SOC")
                else
                    T=DataNEWARE2.(tableNames{i});
                    % Finding the real Capacity curve by integration (the one provided is reset at each phase)
                    Capacity=cumtrapz(DataNEWARE2.(tableNames{i}).("Time(s)"),DataNEWARE2.(tableNames{i}).("Current(A)"))/3600;
                    T.('Real Capacity by Integration(Ah)')=Capacity;
                    
                    % Finding the Capacity corresponding to 0 and 100% SOC
                    rest_idx = strcmp(T.('Step Type'), 'Rest'); % Identify where 'Step Type' is 'Rest'
                    transitions = diff([0; rest_idx; 0]); % Find start and end indices of all continuous blocks
                    starts = find(transitions == 1);
                    ends   = find(transitions == -1) - 1;
                    
                    Capacity0 = T.('Real Capacity by Integration(Ah)')(starts(3)); % Retrieve the corresponding 'Real Capacity by Integration (Ah)'
                    Capacity100 = T.('Real Capacity by Integration(Ah)')(ends(2)); % Retrieve the corresponding 'Real Capacity by Integration (Ah)'
                    
                    % Finding maximum capacity during the pilot discharge phase (used as reference)
                    maxCapacity=Capacity100-Capacity0;
                    % Finding SOC
                    T.('SOC')=(T.('Real Capacity by Integration(Ah)')-Capacity0)./maxCapacity.*100;
                    
                    % Delete useless coloumns and rename some of them for more clarity
                    T(:, {'DataPoint', 'Time','Total Time','Capacity(Ah)','Energy(Wh)','Date','Power(W)'}) = [];
                    T.Properties.VariableNames{strcmp(T.Properties.VariableNames, 'Real Capacity by Integration(Ah)')} = 'Capacity(Ah)';

                    DataNEWARE2.(tableNames{i})=T;
                    DataNEWARE2.(tableNames{i}).('Discharge capacity (Ah)')(:)=maxCapacity; % Save value
                end
            end

            % Calculate Heat Flux from Voltage and automatically add it to data
            Sensitivity=17.39*10^(-6); % V*m^2/W
            TestArea=12.96*10^(-4); % m^2
            LateralSurfaceArea=2*pi*21*70*10^(-6); % m^2

            tableNames = fieldnames(DataDEWESOFT2); % Get the available table names
            
            % For each of the tests, calculates the heat flux quantities and adds them to the data
            for i = 1:length(tableNames)          
                % Heat fluxes culculation with temperature corrected sensitivity value
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)")=DataDEWESOFT2.(tableNames{i}).("Thermal flux FHF05")./(Sensitivity.*(1+0.002.*(DataDEWESOFT2.(tableNames{i}).("Temperature FHF05")-20))); % W/m^2
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Test Area (W)")=DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)").*TestArea; % W on test area of the sensor
                DataDEWESOFT2.(tableNames{i}).("Heat Flux Lateral Surface Area (W)")=DataDEWESOFT2.(tableNames{i}).("Heat Flux Density (W/m^2)").*LateralSurfaceArea; % W on the lateral surface area

                % Rename 'Thermal flux FHF05' to 'Thermal Flux Sensor Voltage (V)'
                DataDEWESOFT2.(tableNames{i}).Properties.VariableNames{'Thermal flux FHF05'} = 'Thermal Flux Sensor Voltage (V)';

                % Rename 'Temperature FHF05' to 'Temperature (C)'
                DataDEWESOFT2.(tableNames{i}).Properties.VariableNames{'Temperature FHF05'} = 'Temperature (C)';
            end

            % Unite NEWARE and DEWESOFT data
            % Initialize the combined struct (merge of NEWARE and DEWESOFT Data)
            DataCOMBINED = struct();

            % Get the field names of both structs
            fieldsNEWARE = fieldnames(DataNEWARE2);
            fieldsDEWESOFT = fieldnames(DataDEWESOFT2);

            % Loop over the fields of DataNEWARE2 to find corresponding fields in DataDEWESOFT2
            for i = 1:length(fieldsNEWARE)
                % Check if the current field name ends with 'NEWARE'
                if endsWith(fieldsNEWARE{i}, '_NEWARE_xlsx')
                    % Extract the base name by removing 'NEWARE'
                    baseName = extractBefore(fieldsNEWARE{i}, '_NEWARE_xlsx');
                    % Check if the same base name exists in DataDEWESOFT2 with 'DEWESOFT'
                    correspondingField = strcat(baseName, '_DEWESOFT_xlsx');
                    if isfield(DataDEWESOFT2, correspondingField)
                        % Get the tables from both structs
                        tableNEWARE = DataNEWARE2.(fieldsNEWARE{i});
                        tableDEWESOFT = DataDEWESOFT2.(correspondingField);

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

                        % Extract data during the discharge phase
                        DischargeData=filledTable(strcmp(filledTable.("Step Type"), 'CC DChg'), :);
                        DischargeData.('Time(s)')=DischargeData.('Time(s)')-DischargeData.('Time(s)')(1);

                        % Save the combined table in the new struct
                        DataCOMBINED.(baseName) = DischargeData;
                    end
                end
            end

            % Get available table names
            tableNames = fieldnames(DataCOMBINED);
            
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
            selectedTable = DataCOMBINED.(selectedTableName);

            % Export Current
            Current_Profile = timeseries(selectedTable.('Current(A)'), selectedTable.('Time(s)'));
            save('Current_Profile.mat', 'Current_Profile');

            % Export Voltage
            Voltage_Profile = timeseries(selectedTable.('Voltage(V)'), selectedTable.('Time(s)'));
            save('Voltage_Profile.mat', 'Voltage_Profile');

            % Export Temperature
            Temperature_Profile = timeseries(selectedTable.('Temperature (C)'), selectedTable.('Time(s)'));
            save('Temperature_Profile.mat', 'Temperature_Profile');

            % Export Heat Flux Density
            Heat_Flux_Density_Profile = timeseries(selectedTable.('Heat Flux Density (W/m^2)'), selectedTable.('Time(s)'));
            save('Heat_Flux_Density_Profile.mat', 'Heat_Flux_Density_Profile');

            disp('Success: Data exported to .mat files');

        case 15
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

        case 16
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

        case 17
            %% CLOSE
            exit=true;
        otherwise
            exit=true;
    end

end


