function export_figs()
    % --- OPEN FILE SELECTION DIALOG ---
    % Allow user to select multiple .fig files
    [fileNames, filePath] = uigetfile('*.fig', 'Select MATLAB Figures to Convert', 'MultiSelect', 'on');
    
    % Check if user canceled the dialog
    if isequal(fileNames, 0)
        disp('File selection canceled.');
        return;
    end
    
    % If only one file is selected, it returns a string. 
    % We convert it to a cell array so the loop works the same way.
    if ischar(fileNames)
        fileNames = {fileNames};
    end
    
    % --- SETUP EXPORT FOLDER ---
    outputFolder = fullfile(filePath, 'png_exports');
    if ~exist(outputFolder, 'dir'); mkdir(outputFolder); end
    
    % --- STYLE SETTINGS ---
    figWidth = 10;   
    figHeight = 8;   
    
    baseTextFS = 20; 
    labelFS = 24;    
    tickFS = 18;     
    legendFS = 20;   
    
    dataLineWidth = 2;    
    legendBoxWidth = 0.5; 

    fprintf('Processing %d selected files...\n', length(fileNames));

    % --- LOOP THROUGH SELECTED FILES ---
    for i = 1:length(fileNames)
        filename = fileNames{i};
        fullFilePath = fullfile(filePath, filename);
        
        fprintf('Processing: %s\n', filename);
        figHandle = openfig(fullFilePath, 'invisible');
        
        % ============================================================
        % STEP 1: CLEANUP & PREP
        % ============================================================
        set(figHandle, 'Color', 'w');        
        set(figHandle, 'InvertHardcopy', 'off'); 
        
        allAxes = findall(figHandle, 'type', 'axes');
        for a = 1:length(allAxes)
            if isprop(allAxes(a), 'Title')
                set(allAxes(a).Title, 'String', '');
                set(allAxes(a).Title, 'Visible', 'off'); 
            end
        end
        
        allTiles = findall(figHandle, 'type', 'tiledlayout');
        for t = 1:length(allTiles)
            title(allTiles(t), ''); 
        end

        % ============================================================
        % STEP 2: RESIZE ALL TEXT (The "Units" Fix)
        % ============================================================
        allTextObjs = findall(figHandle, '-property', 'FontSize');
        for t = 1:length(allTextObjs)
            obj = allTextObjs(t);
            if isprop(obj, 'FontUnits'); set(obj, 'FontUnits', 'points'); end
            set(obj, 'FontSize', baseTextFS);
            if isprop(obj, 'Color'); set(obj, 'Color', 'k'); end
            if isprop(obj, 'TextColor'); set(obj, 'TextColor', 'k'); end
        end
        
        % ============================================================
        % STEP 3: FINE-TUNE AXES
        % ============================================================
        for a = 1:length(allAxes)
            ax = allAxes(a);
            set(ax, 'FontSize', tickFS); 
            set(ax, 'Color', 'none'); 
            
            if isprop(ax, 'XColor') && mean(ax.XColor) > 0.80; ax.XColor = 'k'; end
            
            if isprop(ax, 'XLabel')
                set(ax.XLabel, 'FontSize', labelFS);
                set(ax.XLabel, 'Color', 'k'); 
            end
            if isprop(ax, 'YLabel'); set(ax.YLabel, 'FontSize', labelFS); end
            if isprop(ax, 'ZLabel'); set(ax.ZLabel, 'FontSize', labelFS); end

            for y = 1:numel(ax.YAxis)
                yAx = ax.YAxis(y);
                if mean(yAx.Color) > 0.80
                     if mean(yAx.Color) > 0.95; yAx.Color = 'k';
                     else; yAx.Color = yAx.Color * 0.6; end
                end
                yAx.FontSize = tickFS;        
                yAx.Label.FontSize = labelFS; 
                yAx.Label.Color = yAx.Color;
            end
        end
        
        % ============================================================
        % STEP 4: THICKEN LINES
        % ============================================================
        plotObjs = findall(figHandle, '-property', 'LineWidth');
        for p = 1:length(plotObjs)
            obj = plotObjs(p);
            if strcmpi(obj.Type, 'axes') || strcmpi(obj.Type, 'legend'); continue; end
            
            set(obj, 'LineWidth', dataLineWidth);
            if isprop(obj, 'Color')
                col = get(obj, 'Color');
                if isnumeric(col) && length(col) == 3 && mean(col) > 0.70
                   if mean(col) > 0.95; set(obj, 'Color', 'k');
                   else; set(obj, 'Color', col * 0.6); end
                end
            end
        end
        
        % ============================================================
        % STEP 5: LEGEND CLEANUP
        % ============================================================
        legs = findall(figHandle, 'type', 'legend');
        for lg = 1:length(legs)
            set(legs(lg), 'FontUnits', 'points');
            set(legs(lg), 'FontSize', legendFS); 
            set(legs(lg), 'LineWidth', legendBoxWidth); 
            set(legs(lg), 'Color', 'w');
            set(legs(lg), 'TextColor', 'k');
        end

        % ============================================================
        % STEP 6: EXPORT
        % ============================================================
        set(figHandle, 'PaperUnits', 'inches');
        set(figHandle, 'PaperPosition', [0 0 figWidth figHeight]);
        set(figHandle, 'PaperPositionMode', 'manual');
        set(figHandle, 'PaperSize', [figWidth figHeight]);
        
        [~, name, ~] = fileparts(filename);
        outputName = fullfile(outputFolder, [name '.png']);
        
        print(figHandle, outputName, '-dpng', '-r300');
        close(figHandle);
    end
    disp('Selected files successfully exported!');
end