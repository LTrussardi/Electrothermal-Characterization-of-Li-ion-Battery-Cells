function dataStruct = ImportFiles(sheetName, varName, optionalSheetName)
    % Check if the variable already exists in the workspace
    if evalin('base', sprintf('exist(''%s'', ''var'')', varName))
        % Load existing data structure from the workspace
        dataStruct = evalin('base', varName);
    else
        % Initialize an empty structure if it does not exist
        dataStruct = struct();
    end

    % Open file selection dialog for Excel files
    [fileNames, folderPath] = uigetfile({'*.xls*', 'Excel Files (*.xls, *.xlsx, *.xlsm)'}, ...
                                        'Select Excel Files to Import', ...
                                        'MultiSelect', 'on');

    % Check if the user canceled the selection
    if isequal(fileNames, 0)
        fprintf('No files selected. Exiting function.\n');
        return;
    end

    % Ensure fileNames is a cell array (handles single-file selection case)
    if ischar(fileNames)
        fileNames = {fileNames};
    end

    % Loop through each selected Excel file
    for i = 1:length(fileNames)
        fileName = fileNames{i};
        filePath = fullfile(folderPath, fileName);
        
        try
            % Import the primary sheet from the Excel file
            data = readtable(filePath, 'Sheet', sheetName, 'PreserveVariableNames', true);
            
            % If an optional sheet name is provided, check if it exists
            if nargin == 3 && ~isempty(optionalSheetName)
                [~, sheets] = xlsfinfo(filePath); % Get list of sheet names
                
                if any(strcmp(sheets, optionalSheetName))
                    % Import the additional sheet
                    extraData = readtable(filePath, 'Sheet', optionalSheetName, 'PreserveVariableNames', true);
                    
                    % Concatenate both sheets (assuming they have the same columns)
                    data = [data; extraData]; 
                else
                    %fprintf('Sheet "%s" not found in %s. Skipping concatenation.\n', optionalSheetName, fileName);
                end
            end
            
            % Create a valid field name
            sanitizedFileName = matlab.lang.makeValidName(fileName);
            
            % Store the concatenated data in the structure
            dataStruct.(sanitizedFileName) = data;
            
            if nargin == 3 && any(strcmp(sheets, optionalSheetName))
                sheetInfo = [' + ' optionalSheetName];
            else
                sheetInfo = '';
            end
            
            fprintf('Imported: %s (Sheets: %s%s)\n', fileName, sheetName, sheetInfo);

        catch ME
            fprintf('Error importing %s: %s\n', fileName, ME.message);
        end
    end
end
