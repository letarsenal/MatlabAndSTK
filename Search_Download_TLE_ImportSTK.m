function Search_Download_TLE_ImportSTK
% MATLAB_STK_TLE_Importer
% GUI to download TLEs from CelesTrak by NORAD CAT IDs and import into STK.
% Integrated: SATCAT JSON fuzzy search, table-select bulk TLE download, STK import.
% Added: UI progress dialog showing stages: 下载TLE开始 -> 下载完成 -> 连接STK -> 建立场景 -> 导入TLE完成
%
% Save as MATLAB_STK_TLE_Importer_with_progress.m and run in MATLAB on Windows with STK installed.

    % 字体设置 - 使用微软雅黑，更清晰醒目
    fontName = 'Microsoft YaHei';  % 微软雅黑，Windows上广泛支持的中文字体
    titleFontSize = 14;           % 标题字体大小
    labelFontSize = 14;           % 标签字体大小
    buttonFontSize = 14;          % 按钮字体大小
    tableFontSize = 11;           % 表格字体大小
    statusFontSize = 11;          % 状态栏字体大小

    defaultGroundElevation = 0; % km

    % Program directory (script location or pwd)
    if ismcc || isdeployed
        programDir = pwd;
    else
        [scriptPath, ~, ~] = fileparts(mfilename('fullpath'));
        if isempty(scriptPath)
            programDir = pwd;
        else
            programDir = scriptPath;
        end
    end

    % default files
    app.tleTempFile = fullfile(programDir, 'downloaded_satellites.tle');
    app.resultsDir = fullfile(programDir, 'CZML_Results');

    % Create UIFigure (use uifigure for modern layout)
    app.fig = uifigure('Name','在轨航天器TLE查询下载及三维演示软件','Position',[200 200 1200 800],...
        'Color', [0.96 0.96 0.96]); % 添加浅灰色背景

    % create placeholder for progress dialog reference
    app.progressDlg = [];

    % 增加窗口高度以适应顶部标题
    leftX = 20; leftW = 620; rightX = 660; rightW = 320;
    topY = 680; % 调整topY位置，为顶部标题留出空间

% ============== 软件标题（居中） ==============
    % 计算标题位置，使其居中
    titleWidth = 800;
    titleHeight = 40;
    titleX = (1200 - titleWidth) / 2;

    % 创建标题标签
    titleLabel = uilabel(app.fig, 'Position', [titleX topY+50 titleWidth titleHeight], ...
    'Text', '在轨航天器TLE查询下载及三维演示软件', ...
    'FontWeight', 'bold', 'FontSize', 24, 'FontName', fontName, ...
    'FontColor', [0.1 0.3 0.8], 'HorizontalAlignment', 'center', ...
    'VerticalAlignment', 'center', 'BackgroundColor', [0.9 0.95 1]);

    % ============== Left column: Search & Results ==============
    uilabel(app.fig,'Position',[leftX topY 260 24],'Text','卫星搜索 (CelesTrak)', ...
        'FontWeight','bold','FontSize',titleFontSize,'FontName',fontName,...
        'FontColor', [0.1 0.3 0.6]); % 深蓝色标题

    % Name box (exact-ish)
    uilabel(app.fig,'Position',[leftX topY-40 70 22],'Text','名称:',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);
    app.nameBox = uieditfield(app.fig,'text','Position',[leftX+70 topY-42 200 26],...
        'Value','','FontSize',labelFontSize,'FontName',fontName);

    % NORAD
    uilabel(app.fig,'Position',[leftX topY-70 150 22],'Text','NORAD编号:',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);
    app.noradBox = uieditfield(app.fig,'text','Position',[leftX+150 topY-72 160 26],...
        'Value','','FontSize',labelFontSize,'FontName',fontName);

    % Max responses
    uilabel(app.fig,'Position',[leftX+300 topY-40 100 22],'Text','最大返回数:',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);
    app.maxResponsesBox = uieditfield(app.fig,'numeric','Position',[leftX+400 topY-42 80 26], ...
        'Value', 500, 'Limits', [1 5000], 'FontSize',labelFontSize,'FontName',fontName);

    % Filters
    uilabel(app.fig,'Position',[leftX topY-100 140 22],'Text','筛选条件:',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);
    app.chkPayloads = uicheckbox(app.fig,'Text','有效载荷','Position',[leftX+110 topY-102 100 22],...
        'Value',true,'FontSize',labelFontSize,'FontName',fontName,'FontWeight','bold');
    app.chkActive   = uicheckbox(app.fig,'Text','活跃','Position',[leftX+220 topY-102 100 22],...
        'Value',true,'FontSize',labelFontSize,'FontName',fontName,'FontWeight','bold');
    app.chkOnOrbit  = uicheckbox(app.fig,'Text','在轨','Position',[leftX+330 topY-102 100 22],...
        'Value',true,'FontSize',labelFontSize,'FontName',fontName,'FontWeight','bold');

    % Search button
    app.btnSearch = uibutton(app.fig,'push','Text','搜索卫星','Position',[leftX+460 topY-102 120 40],...
        'ButtonPushedFcn',@(btn,event)searchSatellitesCallback(),...
        'BackgroundColor',[0.2 0.5 0.9],'FontColor',[1 1 1],'FontSize',buttonFontSize,...
        'FontName',fontName,'FontWeight','bold');

    % Results label
    uilabel(app.fig,'Position',[leftX topY-150 200 24],'Text','搜索结果:', ...
        'FontWeight','bold','FontSize',titleFontSize,'FontName',fontName,...
        'FontColor', [0.1 0.3 0.6]);

    % Table column names: add first column "Select"
    colNames = {'选择','NORAD编号','卫星名称','来源','发射日期','发射场','离轨日期','对象类型','运行状态','最后更新'};

    % initial empty data: show 0 rows but 10 columns (first is logical)
    emptyData = cell(0, numel(colNames));
    % Create uitable
    app.resultTable = uitable(app.fig,'Position',[leftX topY-520 leftW 360],...
        'Data', emptyData, 'ColumnName', colNames, ...
        'ColumnEditable', [true false false false false false false false false false], ...
        'CellEditCallback', @(tbl,event) resultTableCellEditCallback(tbl,event),...
        'FontSize',tableFontSize,'FontName',fontName,...
        'RowStriping','on','BackgroundColor',[1 1 1]);

    % Place select-all checkbox above the "Select" column header
    % We'll align it visually; compute a reasonable position
    uilabel(app.fig,'Position',[leftX+10 topY-540 60 22],'Text','全选:',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);

    headerCheckboxX = leftX + 60;
    headerCheckboxY = topY-540;
    app.chkSelectAll = uicheckbox(app.fig,'Text','','Position',[headerCheckboxX headerCheckboxY 20 20],...
        'Value',false,'ValueChangedFcn',@(chk,event) selectAllChanged(chk));

    % Status log (left bottom)
    app.status = uilabel(app.fig,'Position',[leftX 10 leftW 24],'Text','就绪',...
        'FontAngle','italic','FontSize',statusFontSize,'FontName',fontName,...
        'FontWeight','bold','FontColor',[0 0.4 0]);

    % ============== Right column: TLE Download / STK Import / Options ==============
    uilabel(app.fig,'Position',[rightX+100 topY-10 300 24],'Text','TLE 下载 / STK 导入 和 场景设置', ...
        'FontWeight','bold','FontSize',titleFontSize,'FontName',fontName,...
        'FontColor', [0.6 0.1 0.1]); % 深红色标题

    % NORAD input box area (textarea) - also will be auto-filled by selection
    uilabel(app.fig,'Position',[rightX+100 topY-50 300 20],'Text','要下载的 NORAD 编号（逗号分隔）:',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);
    app.tleBox = uitextarea(app.fig,'Position',[rightX+100 topY-160 300 100],...
        'Value',{'56727,56728'}, 'FontSize',labelFontSize,'FontName',fontName);

    % Buttons (Download / Import)
    app.btnDownload = uibutton(app.fig,'push','Text','下载并保存 TLE','Position',[rightX+100 topY-210 140 36],...
        'ButtonPushedFcn',@(btn,event)downloadTLECallback(),...
        'BackgroundColor',[0.2 0.5 0.9],'FontColor',[1 1 1],'FontSize',buttonFontSize,...
        'FontName',fontName,'FontWeight','bold');
    
    % ============== 新增：下载进度条区域 ==============
    uilabel(app.fig,'Position',[rightX+100 topY-250 80 22],'Text','下载进度:',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);
    
    % 进度条标签（显示百分比）
    app.progressBarLabel = uilabel(app.fig,'Position',[rightX+185 topY-250 120 22],...
        'Text','0% (0/0)','FontSize',labelFontSize,'FontName',fontName,'FontWeight','bold');
    
    % 使用uipanel模拟进度条（使用uilabel作为边框）
    progressBarBgPos = [rightX+310 topY-250 90 22];
    
    % 创建一个有边框的背景标签作为进度条背景
    app.progressBarBg = uilabel(app.fig,'Position',progressBarBgPos,...
        'BackgroundColor',[0.9 0.9 0.9]); % 设置小字体，使其不显示文本
    
    % 创建一个uilabel作为边框（模拟uipanel的边框效果）
    app.progressBarBorder = uilabel(app.fig,'Position',[progressBarBgPos(1)-1, progressBarBgPos(2)-1, progressBarBgPos(3)+2, progressBarBgPos(4)+2],...
        'BackgroundColor',[0.7 0.7 0.7]);
    
    % 进度条前景（实际进度显示）- 使用uilabel而不是uipanel
    app.progressBarFg = uilabel(app.fig,'Position',[progressBarBgPos(1) progressBarBgPos(2) 0 progressBarBgPos(4)],...
        'BackgroundColor',[0.2 0.6 1.0]);
    
    % 进度条数值显示标签（在进度条上方显示当前下载的NORAD编号）
    app.progressDetailLabel = uilabel(app.fig,'Position',[rightX+100 topY-280 300 24],...
        'Text','准备下载...','FontSize',11,'FontName',fontName,...
        'HorizontalAlignment','left','FontColor',[0.2 0.2 0.6]);

    % Time range
    uilabel(app.fig,'Position',[rightX+100 topY-310 150 22],'Text','场景时间范围(天):',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);
    app.timeRangeEdit = uieditfield(app.fig, 'numeric','Position',[rightX+250 topY-310 80 26],...
        'Value',1,'Limits',[0.1 30],'FontSize',labelFontSize,'FontName',fontName);

    % Ground station selection (multi-select)
    uilabel(app.fig,'Position',[rightX+100 topY-360 300 22],'Text','设置地面站（可多选）:',...
        'FontWeight','bold','FontSize',labelFontSize,'FontName',fontName);
    app.groundStationList = uilistbox(app.fig,'Position',[rightX+100 topY-510 300 150],...
        'Items', {'北京', '宁波', '喀什', '三亚', '七台河'}, 'Multiselect','on', ...
        'Value', {'北京'},'FontSize',labelFontSize,'FontName',fontName);

    % Add cities checkbox
    app.chkAddCities = uicheckbox(app.fig,'Text','中国省会城市作为地面站添加到场景',...
        'Position',[rightX+100 topY-540 280 22],'Value',false,...
        'FontSize',labelFontSize,'FontName',fontName,'FontWeight','bold');
    
    app.btnToSTK = uibutton(app.fig,'push','Text','连接STK软件并导入TLE中所有卫星','Position',[rightX+100 topY-600 300 40],...
    'ButtonPushedFcn',@(btn,event)importToSTKCallback(),...
    'BackgroundColor',[0.2 0.5 0.9],'FontColor',[1 1 1],'FontSize',buttonFontSize,...
    'FontName',fontName,'FontWeight','bold');

    % Author label
    uilabel(app.fig,'Position',[rightX+450 topY-650 100 30],'Text','作者：梁尔涛',...
        'FontSize',11,'FontName',fontName,'FontWeight','bold','FontColor',[0.4 0.4 0.4]);
    uilabel(app.fig,'Position',[rightX+400 topY-680 150 30],'Text','huoyan_sat@163.com',...
    'FontSize',11,'FontName',fontName,'FontWeight','bold','FontColor',[0.4 0.4 0.4]);

    % 创建分隔线的替代方法：使用一个很窄的uipanel作为分隔线
    % 这里使用uilabel替代uipanel作为分隔线
    separatorX = rightX - 10; % 分隔线位置
    separatorWidth = 2; % 分隔线宽度
    uilabel(app.fig,'Position',[separatorX, 20, separatorWidth, 660],...
        'BackgroundColor',[0.7 0.7 0.7]);

    % Save app into figure UserData to keep lifetime
    app.fig.UserData = app;

    % ----------------- Callback and helper functions -----------------
    % ====== Progress helper (新增) ======
    function updateProgress(value, msg)
        % value: 0..1, msg: string
        try
            % ensure app reference
            app = app.fig.UserData;
            if isempty(app.progressDlg) || ~isvalid(app.progressDlg)
                % create a uiprogressdlg attached to UIFigure
                try
                    app.progressDlg = uiprogressdlg(app.fig, 'Title', '任务进度', 'Message', msg, ...
                        'Cancelable', false, 'FontName', fontName, 'FontSize', labelFontSize);
                catch
                    % older MATLAB may not support uiprogressdlg with UIFigure; try without parent
                    try
                        app.progressDlg = uiprogressdlg('Title', '任务进度', 'Message', msg, ...
                            'Cancelable', false, 'FontName', fontName, 'FontSize', labelFontSize);
                    catch
                        app.progressDlg = [];
                    end
                end
            end
            if ~isempty(app.progressDlg) && isvalid(app.progressDlg)
                if nargin >= 1 && ~isempty(value)
                    try
                        app.progressDlg.Value = min(max(value,0),1);
                    catch
                    end
                end
                if nargin >= 2 && ~isempty(msg)
                    try
                        app.progressDlg.Message = msg;
                    catch
                    end
                end
                % store back
                app.fig.UserData = app;
            end
            drawnow;
            % auto-close when finished
            if ~isempty(app.progressDlg) && isvalid(app.progressDlg) && value >= 1
                try close(app.progressDlg); end
                app.progressDlg = [];
                app.fig.UserData = app;
            end
        catch
            % ignore progress failures
        end
    end

    %% ===== 替换：searchSatellitesCallback （使用 CelesTrak SATCAT JSON 模糊搜索） =====
%% ===== 替换：searchSatellitesCallback （修复不同API响应格式问题） =====
    function searchSatellitesCallback()
        % refresh app reference
        app = app.fig.UserData;

        % 定义表格列名
        colNames = {'Select', 'NORAD ID', 'Name', 'Source', 'Launch Date', ...
                    'Launch Site', 'Decay Date', 'Type', 'Status', 'Last Update'};

        % Get UI values
        nameTerm = strtrim(app.nameBox.Value);
        noradTerm = strtrim(app.noradBox.Value);
        maxResponses = app.maxResponsesBox.Value;

        % filters
        payloadsFilter = app.chkPayloads.Value;
        activeFilter = app.chkActive.Value;
        onOrbitFilter = app.chkOnOrbit.Value;

        app.status.Text = '正在使用 CelesTrak SATCAT API 搜索...';
        app.status.FontColor = [0 0.4 0]; % 绿色状态
        drawnow;

        try
            % 优先处理NORAD编号搜索
            if ~isempty(noradTerm)
                % 使用返回JSON数据的API端点 - 修正参数名为CATNR
                url = sprintf('https://celestrak.org/satcat/record.php?CATNR=%s&FORMAT=JSON', noradTerm);

                % 查询卫星记录
                try
                    satRecord = webread(url);

                    % 处理返回结果
                    if isempty(satRecord) || (isstruct(satRecord) && isfield(satRecord, 'error'))
                        app.resultTable.Data = cell(0,numel(colNames));
                        app.status.Text = sprintf('未找到NORAD编号为 %s 的卫星', noradTerm);
                        app.status.FontColor = [0.8 0 0]; % 红色警告
                        return;
                    end

                    % 将结果转换为元胞数组 - NORAD搜索返回单个记录
                    if ~iscell(satRecord)
                        satlist = {satRecord};
                    else
                        satlist = satRecord;
                    end
                catch ME
                    app.resultTable.Data = cell(0,numel(colNames));
                    app.status.Text = ['查询卫星记录失败: ' ME.message];
                    app.status.FontColor = [0.8 0 0]; % 红色错误
                    return;
                end
            else
                % 名称模糊搜索逻辑保持不变
                optsMax = maxResponses;
                activeParam = '';
                onOrbitParam = '';
                payloadsParam = '';
                if activeFilter, activeParam = '&ACTIVE=1'; end
                if onOrbitFilter, onOrbitParam = '&ONORBIT=1'; end
                if payloadsFilter, payloadsParam = '&PAYLOADS=1'; end

                url = sprintf('https://celestrak.org/satcat/records.php?NAME=%s&FORMAT=JSON&MAX=%d%s%s%s', ...
                    nameTerm, optsMax, activeParam, onOrbitParam, payloadsParam);

                % 查询CelesTrak SATCAT
                satlist = webread(url);
            end

            % 确保satlist是元胞数组
            if ~iscell(satlist)
                if isstruct(satlist) && numel(satlist) == 1
                    satlist = {satlist};
                else
                    satlist = num2cell(satlist);
                end
            end

            if isempty(satlist)
                app.resultTable.Data = cell(0,numel(colNames));
                app.status.Text = '未找到匹配的卫星（CelesTrak 返回空或请求失败）';
                app.status.FontColor = [0.8 0 0]; % 红色警告
                return;
            end

            % 准备表格数据 - 根据搜索类型使用不同的字段映射
            N = numel(satlist);
            tbl = cell(N,numel(colNames));

            for i=1:N
                % 安全获取当前卫星记录
                if iscell(satlist)
                    s = satlist{i};
                else
                    s = satlist(i);
                end

                % 确保s是结构体
                if ~isstruct(s)
                    continue;
                end

                % 判断是哪种API响应格式
                hasCatNumber = isfield(s, 'catNumber');  % NORAD搜索格式标识
                hasObjectName = isfield(s, 'OBJECT_NAME'); % 名称搜索格式标识

                if hasCatNumber
                    % ===== NORAD编号搜索结果处理（图1格式）=====
                    catnr = getFieldSafe(s, 'catNumber', '');
                    oname = getFieldSafe(s, 'satName', '');
                    source = 'CelesTrak';
                    ldate = getFieldSafe(s, 'launchDate', '');
                    lsite = getFieldSafe(s, 'launchSite', '');
                    ddate = getFieldSafe(s, 'decayDate', '');
                    otype = ''; % NORAD搜索格式似乎没有直接的object_type字段
                    statusf = getFieldSafe(s, 'opsStatus', '');
                    lastupd = datestr(now,'yyyy-mm-dd'); % 该格式没有last_update字段

                elseif hasObjectName
                    % ===== 名称搜索结果处理（图2格式）=====
                    catnr = getFieldSafe(s, 'NORAD_CAT_ID', '');
                    oname = getFieldSafe(s, 'OBJECT_NAME', '');
                    source = 'CelesTrak';
                    ldate = getFieldSafe(s, 'LAUNCH_DATE', '');
                    lsite = getFieldSafe(s, 'LAUNCH_SITE', '');
                    ddate = getFieldSafe(s, 'DECAY_DATE', '');
                    otype = getFieldSafe(s, 'OBJECT_TYPE', '');
                    statusf = getFieldSafe(s, 'OPS_STATUS_CODE', '');
                    lastupd = datestr(now,'yyyy-mm-dd'); % 该格式也没有last_update字段

                else
                    % ===== 兼容旧格式或其他格式 =====
                    catnr = getFieldSafe(s, 'CATNR', '');
                    if isempty(catnr), catnr = getFieldSafe(s, 'NORAD_CAT_ID', ''); end
                    if isempty(catnr), catnr = getFieldSafe(s, 'NORAD_CAT_NR', ''); end
                    if isempty(catnr), catnr = getFieldSafe(s, 'NORAD_ID', ''); end
                    oname = getFieldSafe(s, 'OBJECT_NAME', '');
                    if isempty(oname), oname = getFieldSafe(s, 'satName', ''); end
                    if isempty(oname), oname = getFieldSafe(s, 'OBJECT_NAME', ''); end
                    source = getFieldSafe(s, 'SOURCE_CODE', getFieldSafe(s, 'SOURCE', 'CelesTrak'));
                    ldate = getFieldSafe(s, 'LAUNCH_DATE', '');
                    if isempty(ldate), ldate = getFieldSafe(s, 'launchDate', ''); end
                    lsite = getFieldSafe(s, 'LAUNCH_SITE', '');
                    if isempty(lsite), lsite = getFieldSafe(s, 'launchSite', ''); end
                    ddate = getFieldSafe(s, 'DECAY_DATE', '');
                    if isempty(ddate), ddate = getFieldSafe(s, 'decayDate', ''); end
                    otype = getFieldSafe(s, 'OBJECT_TYPE', '');
                    statusf = getFieldSafe(s, 'STATUS', getFieldSafe(s, 'OPS_STATUS', ''));
                    if isempty(statusf), statusf = getFieldSafe(s, 'OPS_STATUS_CODE', ''); end
                    lastupd = getFieldSafe(s, 'LAST_UPDATE', datestr(now,'yyyy-mm-dd'));
                end

                tbl{i,1} = false;            % Select checkbox (default false)
                tbl{i,2} = catnr;
                tbl{i,3} = oname;
                tbl{i,4} = source;
                tbl{i,5} = ldate;
                tbl{i,6} = lsite;
                tbl{i,7} = ddate;
                tbl{i,8} = otype;
                tbl{i,9} = statusf;
                tbl{i,10} = lastupd;
            end

            % 设置表格数据
            app.resultTable.Data = tbl;
            app.resultTable.ColumnName = colNames;  % 确保列名设置

            if ~isempty(noradTerm)
                app.status.Text = sprintf('成功查询NORAD编号 %s 对应的卫星', noradTerm);
            else
                app.status.Text = sprintf('搜索完成，共返回 %d 条结果', N);
            end
            app.status.FontColor = [0 0.4 0]; % 绿色成功
            % ensure header checkbox unchecked after new results
            app.chkSelectAll.Value = false;
            app.fig.UserData = app;
        catch err
            app.resultTable.Data = cell(0,numel(colNames));
            app.status.Text = ['搜索失败: ' err.message];
            app.status.FontColor = [0.8 0 0]; % 红色错误
        end
        drawnow;
    end

    function value = getFieldSafe(s, field, default)
        % 安全获取结构体字段值
        if isfield(s, field)
            fieldValue = s.(field);
            % 处理字符串字段可能的额外引号
            if ischar(fieldValue) && length(fieldValue) >= 2 && ...
               fieldValue(1) == '''' 
                value = fieldValue(2:end-1);
            else
                value = fieldValue;
            end
        else
            value = default;
        end
    end


    %% ===== 新增：resultTableCellEditCallback =====
    function resultTableCellEditCallback(tbl,event)
        % event.Indices = [row, col], event.NewData, event.PreviousData
        % We care when column 1 ("Select") is edited.
        app = app.fig.UserData;
        try
            if isempty(event.Indices), return; end
            row = event.Indices(1);
            col = event.Indices(2);
            if col == 1
                % Update NORAD text area based on all selected rows
                data = app.resultTable.Data;
                if isempty(data)
                    app.tleBox.Value = {''};
                    return;
                end
                selectedIds = {};
                for r = 1:size(data,1)
                    sel = data{r,1};
                    if isequal(sel,true)
                        idVal = data{r,2};
                        if ~isempty(idVal)
                            selectedIds{end+1} = num2str(idVal); %#ok<AGROW>
                        end
                    end
                end
                if isempty(selectedIds)
                    % keep existing TLE box content? According to spec, clear it
                    app.tleBox.Value = {''};
                else
                    % join with commas and set as single string inside cell array (uitextarea uses cellstr or string)
                    joined = strjoin(selectedIds, ',');
                    app.tleBox.Value = {joined};
                end
                % if any were unchecked, uncheck select-all
                allSelected = all(cellfun(@(x) isequal(x,true), data(:,1)));
                app.chkSelectAll.Value = allSelected;
                app.fig.UserData = app;
            end
        catch err
            warning('处理表格编辑回调时出错: %s', err.message);
        end
    end

    %% ===== 新增：selectAllChanged =====
    function selectAllChanged(chk)
        app = app.fig.UserData;
        try
            if isempty(app.resultTable.Data), return; end
            data = app.resultTable.Data;
            n = size(data,1);
            for i=1:n
                data{i,1} = chk.Value; % set checkbox column
            end
            app.resultTable.Data = data; % trigger UI update
            % update TLE box accordingly
            if chk.Value
                % all selected: collect all NORADs
                ids = cell(0,1);
                for r=1:n
                    idVal = data{r,2};
                    if ~isempty(idVal)
                        ids{end+1} = num2str(idVal); %#ok<AGROW>
                    end
                end
                if isempty(ids)
                    app.tleBox.Value = {''};
                else
                    app.tleBox.Value = {strjoin(ids,',')};
                end
            else
                % unselect all
                app.tleBox.Value = {''};
            end
            app.fig.UserData = app;
        catch err
            warning('处理全选复选框时出错: %s', err.message);
        end
    end

%% ===== 新增：更新进度条函数 =====
function updateProgressBar(progressPercent, currentItem, totalItems, currentNORAD)
    % progressPercent: 0-100 的进度百分比
    % currentItem: 当前处理的第几个项目
    % totalItems: 总项目数
    % currentNORAD: 当前正在下载的NORAD编号
    
    app = app.fig.UserData;
    
    try
        % 更新进度条标签 - 只显示百分比，不显示计数
        app.progressBarLabel.Text = sprintf('%.1f%%', progressPercent);
        

        % 更新详细标签
        if nargin >= 4 && ~isempty(currentNORAD)
            app.progressDetailLabel.Text = sprintf('正在下载 NORAD: %s %d/%d', currentNORAD, currentItem, totalItems);
        elseif currentItem > 0 && totalItems > 0
            app.progressDetailLabel.Text = sprintf('正在处理第 %d/%d 项', currentItem, totalItems);
        end
        
         if isfield(app,'progressBarBg') && isfield(app,'progressBarFg') && ...
           isvalid(app.progressBarBg) && isvalid(app.progressBarFg)

            bgPos = app.progressBarBg.Position;
            newWidth = bgPos(3) * progressPercent / 100;

            app.progressBarFg.Position = ...
                [bgPos(1), bgPos(2), max(1,newWidth), bgPos(4)];
        end
        drawnow;
    catch err
        warning('更新进度条时出错: %s', err.message);
    end
end

%% ===== 新增：downloadTLEsForIDs（包含进度回调） =====
function count = downloadTLEsForIDs(idList, outFile)
    count = 0;
    fid = fopen(outFile,'w');
    if fid == -1
        error('无法创建或写入文件: %s', outFile);
    end

    nTotal = numel(idList);
    app = app.fig.UserData;
    
    % 检测是否为部署环境
    isDeployed = isdeployed || ~isempty(ver('compiler'));
    
    for k = 1:nTotal
        id = idList{k};
        id = strtrim(id);
                
        % 更新进度条 - 添加这一部分
        progressPercent = k/nTotal * 100;
        updateProgressBar(progressPercent, k, nTotal, id);
        drawnow; % 强制刷新UI
        
        if isempty(id)
            continue;
        end
        
        try
            url = sprintf('https://celestrak.org/NORAD/elements/gp.php?CATNR=%s&FORMAT=TLE', id);
            
            % 针对部署环境使用不同的下载方法
            if isDeployed
                % 方法1：使用Java网络库
                try
                    import java.net.*
                    import java.io.*
                    
                    urlObj = URL(url);
                    connection = urlObj.openConnection();
                    connection.setRequestMethod('GET');
                    connection.setRequestProperty('User-Agent', 'Mozilla/5.0');
                    connection.setConnectTimeout(10000);
                    connection.setReadTimeout(10000);
                    
                    inputStream = connection.getInputStream();
                    reader = BufferedReader(InputStreamReader(inputStream));
                    
                    line = reader.readLine();
                    tleText = '';
                    while ~isempty(line)
                        tleText = [tleText char(line) newline]; %#ok<AGROW>
                        line = reader.readLine();
                    end
                    
                    reader.close();
                    inputStream.close();
                catch javaErr
                    warning('Java下载失败: %s', javaErr.message);
                    tleText = '';
                end
                
                % 方法2：使用系统curl命令
                if isempty(tleText)
                    try
                        if ispc
                            curlCmd = ['curl -s -k -L "' url '"'];
                        else
                            curlCmd = ['curl -s -k -L "' url '"'];
                        end
                        [status, response] = system(curlCmd);
                        
                        if status == 0 && ~isempty(response)
                            tleText = response;
                        end
                    catch curlErr
                        warning('CURL下载失败: %s', curlErr.message);
                    end
                end
            else
                % 在MATLAB环境中使用webread
                try
                    tleText = webread(url);
                catch
                    tleText = '';
                end
            end
            
            if ~isempty(strtrim(tleText))
                if ~strcmp(tleText(end), newline)
                    tleText = [tleText newline];
                end
                fprintf(fid, '%s', tleText);
                count = count + 1;
            end
        catch innerErr
            warning('下载 NORAD %s 时出错: %s', id, innerErr.message);
        end
    end

    fclose(fid);
end


%% ===== 替换：downloadTLECallback（增强：优先使用表格选择的NORAD批量下载） =====
function downloadTLECallback()
    % refresh app reference
    app = app.fig.UserData;

    % 重置进度条
    updateProgressBar(0, 0, 0, '准备中...');
    app.progressDetailLabel.Text = '准备下载...';
    
    % Try to get selected checkboxes first
    idsToDownload = {};
    try
        tblData = app.resultTable.Data;
        if ~isempty(tblData)
            for r = 1:size(tblData,1)
                try
                    if isequal(tblData{r,1}, true)
                        id = tblData{r,2};
                        if ~isempty(id)
                            idsToDownload{end+1} = num2str(id); %#ok<AGROW>
                        end
                    end
                catch
                    continue;
                end
            end
        end
    catch
        idsToDownload = {};
    end

    % If none selected, parse from text box
    if isempty(idsToDownload)
        idsText = '';
        try
            if iscell(app.tleBox.Value)
                idsText = strjoin(app.tleBox.Value, ' ');
            else
                idsText = char(app.tleBox.Value);
            end
        catch
            idsText = '';
        end
        idsText = strrep(idsText, '，', ',');
        idsParsed = regexp(idsText, '[0-9]+','match');
        idsToDownload = idsParsed;
    end

    if isempty(idsToDownload)
        app.status.Text = '未检测到要下载的 NORAD 编号（请选择表格行或在文本框输入）';
        app.status.FontColor = [0.8 0 0]; % 红色警告
        app.progressDetailLabel.Text = '无卫星需要下载';
        updateProgressBar(0, 0, 0, '无卫星');
        return;
    end

    nTotal = numel(idsToDownload);
    app.status.Text = sprintf('准备下载 %d 颗卫星的 TLE...', nTotal);
    app.status.FontColor = [0 0.4 0]; % 绿色状态
    app.progressDetailLabel.Text = sprintf('准备下载 %d 颗卫星...', nTotal);
    drawnow;

    try
        % show progress: start
        updateProgress(0.05, sprintf('开始下载 TLE... (共 %d 颗卫星)', nTotal));
        
        % 初始化进度条
        updateProgressBar(0, 0, nTotal, '开始下载...');

        tleFile = app.tleTempFile;
        % download and update progress inside function
        downloaded = downloadTLEsForIDs(idsToDownload, tleFile);

        if downloaded > 0
            app.status.Text = sprintf('成功下载 %d 颗卫星的 TLE 并保存到：%s', downloaded, tleFile);
            app.status.FontColor = [0 0.4 0]; % 绿色成功
            updateProgress(0.45, 'TLE 下载完成');
            updateProgressBar(100, nTotal, nTotal, '下载完成');
            app.progressDetailLabel.Text = sprintf('下载完成！成功下载 %d/%d 颗卫星', downloaded, nTotal);
        else
            app.status.Text = '未能下载到任何有效 TLE（请检查编号或网络）';
            app.status.FontColor = [0.8 0 0]; % 红色错误
            updateProgress(0, 'TLE 下载失败');
            updateProgressBar(100, nTotal, nTotal, '下载失败');
            app.progressDetailLabel.Text = '下载完成，但未获取到任何TLE数据';
        end
    catch err
        app.status.Text = ['下载失败: ' err.message];
        app.status.FontColor = [0.8 0 0]; % 红色错误
        updateProgress(0, 'TLE 下载失败');
        updateProgressBar(100, nTotal, nTotal, '下载失败');
        app.progressDetailLabel.Text = sprintf('下载失败: %s', err.message);
    end
    drawnow;
end
    % ============== STK import (same as previous integrated version) ==============
    function importToSTKCallback()
        app.status.Text = '准备导入 STK...';
        app.status.FontColor = [0 0.4 0]; % 绿色状态
        drawnow;

        % refresh app references
        app = app.fig.UserData;

        % read UI fields safely (some controls may not exist in all versions)
        selectedStations = app.groundStationList.Value;
        if isfield(app, 'resultsDirEdit') && isvalid(app.resultsDirEdit)
            resultsDir = strtrim(app.resultsDirEdit.Value);
        else
            resultsDir = app.resultsDir;
        end
        timeRangeDays = app.timeRangeEdit.Value;
        if isfield(app,'chkExportSat')
            exportSat = app.chkExportSat.Value;
        else
            exportSat = false;
        end
        if isfield(app,'chkExportChain')
            exportChain = app.chkExportChain.Value;
        else
            exportChain = false;
        end
        addCities = app.chkAddCities.Value;

        if isempty(resultsDir)
            resultsDir = app.resultsDir;
            if isfield(app, 'resultsDirEdit') && isvalid(app.resultsDirEdit)
                app.resultsDirEdit.Value = resultsDir;
            end
        end
        if ~exist(resultsDir, 'dir')
            try
                mkdir(resultsDir);
            catch
                warning('无法创建目录 %s, 使用程序目录', resultsDir);
                resultsDir = programDir;
                if isfield(app, 'resultsDirEdit') && isvalid(app.resultsDirEdit)
                    app.resultsDirEdit.Value = resultsDir;
                end
            end
        end

        tleTempFile = app.tleTempFile;
        if exist(tleTempFile,'file') ~= 2
            app.status.Text = '未找到 TLE 文件，请先下载并保存。';
            app.status.FontColor = [0.8 0 0]; % 红色警告
            return;
        end

        app.status.Text = '尝试连接 STK...';
        drawnow;
        updateProgress(0.5, '正在连接 STK...');

        try
            [uiap, root] = connectSTK(app);
            updateProgress(0.6, 'STK 连接成功');
        catch ME
            app.status.Text = ['连接STK失败: ' ME.message];
            app.status.FontColor = [0.8 0 0]; % 红色错误
            updateProgress(0, '连接 STK 失败');
            return;
        end

        % create scenario
        try
            scenName = ['TLE_Scenario_' datestr(now,'yyyymmdd_HHMMSS')];
            root.NewScenario(scenName);
            sc = root.CurrentScenario;

            startTime = datestr(now - timeRangeDays/2, 'dd mmm yyyy HH:MM:SS');
            endTime = datestr(now + timeRangeDays/2, 'dd mmm yyyy HH:MM:SS');

            sc.SetTimePeriod(startTime, endTime);
            root.ExecuteCommand('Animate * Reset');

            app.status.Text = sprintf('已创建场景: %s，时间范围: %s 到 %s', scenName, startTime, endTime);
            drawnow;
            updateProgress(0.7, '场景已创建');
        catch ME
            app.status.Text = ['创建场景失败: ' ME.message];
            app.status.FontColor = [0.8 0 0]; % 红色错误
            updateProgress(0, '场景创建失败');
            return;
        end

        % parse TLE file
        try
            tleLines = parseTLEFile(tleTempFile);
            if isempty(tleLines)
                app.status.Text = 'TLE 文件中没有可用数据';
                app.status.FontColor = [0.8 0 0]; % 红色警告
                updateProgress(0, '无可用 TLE');
                return;
            end
        catch e
            app.status.Text = ['解析TLE失败: ' e.message];
            app.status.FontColor = [0.8 0 0]; % 红色错误
            updateProgress(0, '解析 TLE 失败');
            return;
        end

        numSatellites = size(tleLines,1);
        satellitesCreated = 0;
        satelliteNames = {};

        % import satellites with progress update from 0.7 -> 0.95
        for satIndex = 1:numSatellites
            nameLine = tleLines{satIndex,1};
            line1 = tleLines{satIndex,2};
            line2 = tleLines{satIndex,3};

            % sanitize name
            if ~isempty(strtrim(nameLine))
                satelliteName = regexprep(strtrim(nameLine), '[^\w]', '_');
                satelliteName = regexprep(satelliteName, '_+', '_');
                if length(satelliteName) > 30
                    satelliteName = satelliteName(1:30);
                end
            else
                satelliteName = sprintf('Satellite_%d', satIndex);
            end

            originalName = satelliteName;
            nameCounter = 1;
            try
                % ensure unique
                while sc.Children.Contains('eSatellite', satelliteName)
                    satelliteName = sprintf('%s_%d', originalName, nameCounter);
                    nameCounter = nameCounter + 1;
                end
            catch
                % if check fails, just append index
                satelliteName = sprintf('%s_%d', originalName, satIndex);
            end

            % create satellite
            try
                satellite = sc.Children.New('eSatellite', satelliteName);

                % Use STK Connect command to set TLE and propagate
                cmdTLE = sprintf('SetState */Satellite/%s TLE "%s" "%s" TimePeriod "%s" "%s"', satelliteName, line1, line2, startTime, endTime);
                root.ExecuteCommand(cmdTLE);

                satellitesCreated = satellitesCreated + 1;
                satelliteNames{end+1} = satelliteName; %#ok<AGROW>

                % 更新状态栏显示导入进度
                app.status.Text = sprintf('正在导入第 %d/%d 颗卫星: %s', satIndex, numSatellites, satelliteName);
                drawnow;
                
                % update progress
                frac = 0.7 + (satIndex/numSatellites)*0.25;
                updateProgress(frac, sprintf('正在导入卫星 %d/%d...', satIndex, numSatellites));

            catch innerErr
                warning('创建卫星 %s 失败: %s', satelliteName, innerErr.message);
                app.status.Text = sprintf('导入第 %d 颗卫星失败: %s', satIndex, innerErr.message);
                app.status.FontColor = [0.8 0 0]; % 红色错误
                drawnow;
            end
        end

        app.status.Text = sprintf('成功创建 %d/%d 颗卫星', satellitesCreated, numSatellites);
        app.status.FontColor = [0 0.4 0]; % 绿色成功
        drawnow;

        % create constellation and add satellites
        if satellitesCreated > 0
            try
                constName = 'ImportedConstellation';
                root.ExecuteCommand(sprintf('New / */Constellation %s', constName));

                addedCount = 0;
                for i = 1:length(satelliteNames)
                    try
                        root.ExecuteCommand(sprintf('Chains */Constellation/%s Add */Satellite/%s', constName, satelliteNames{i}));
                        addedCount = addedCount + 1;
                        if mod(addedCount, 5) == 0 || i == length(satelliteNames)
                            app.status.Text = sprintf('正在添加卫星到星座 %d/%d...', addedCount, length(satelliteNames));
                            drawnow;
                        end
                    catch innerErr
                        warning('添加卫星 %s 到星座失败: %s', satelliteNames{i}, innerErr.message);
                    end
                end

                app.status.Text = sprintf('星座 %s 已创建，成功添加 %d/%d 颗卫星', constName, addedCount, length(satelliteNames));
                drawnow;
            catch e
                warning('创建星座失败: %s', e.message);
            end
        end

        % create ground stations
        app.status.Text = '创建地面站...';
        drawnow;

        stationInfo = {
            'Beijing', 40.37, 116.83, defaultGroundElevation;
            'Ningbo', 29.29, 121.90, defaultGroundElevation;
            'Kashi', 39.40, 76.00, defaultGroundElevation;
            'Sanya', 18.34, 109.70, defaultGroundElevation;
            'Qitaihe', 45.84, 130.97, defaultGroundElevation
        };
        stationMap = containers.Map({'北京','宁波','喀什','三亚','七台河'}, 1:5);

        createdStations = {};
        for i = 1:length(selectedStations)
            stationNameCN = selectedStations{i};
            if isKey(stationMap, stationNameCN)
                idx = stationMap(stationNameCN);
                engName = stationInfo{idx,1};
                lat = stationInfo{idx,2};
                lon = stationInfo{idx,3};
                alt_km = stationInfo{idx,4};
                try
                    root.ExecuteCommand(sprintf('New / */Facility %s', engName));
                    root.ExecuteCommand(sprintf('SetPosition */Facility/%s Geodetic %f %f UseTerrain', engName, lat, lon));
                    createdStations{end+1} = engName; %#ok<AGROW>
                    
                    % 更新状态栏显示地面站创建进度
                    app.status.Text = sprintf('正在创建地面站 %d/%d: %s', i, length(selectedStations), engName);
                    drawnow;
                catch
                    try
                        facObj = sc.Children.New('eFacility', engName);
                        facObj.Position.AssignGeodetic(lat, lon, alt_km);
                        createdStations{end+1} = engName; %#ok<AGROW>
                        
                        app.status.Text = sprintf('正在创建地面站 %d/%d: %s', i, length(selectedStations), engName);
                        drawnow;
                    catch innerErr
                        warning('创建地面站 %s 失败: %s', engName, innerErr.message);
                    end
                end
            end
        end

        % create chain and compute access
        if ~isempty(createdStations)
            app.status.Text = '创建Chain并计算访问数据...';
            drawnow;
            updateProgress(0.95, '创建 Chain 并计算访问...');
            try
                chainName = 'Chain_LEO_GroundStation';
                root.ExecuteCommand(sprintf('New / */Chain %s', chainName));
                root.ExecuteCommand(sprintf('Chains */Chain/%s Add */Facility/%s', chainName, createdStations{1}));
                root.ExecuteCommand(sprintf('Chains */Chain/%s Add */Constellation/%s', chainName, constName));
                app.status.Text = 'Chain已创建，正在计算访问数据...';
                drawnow;

                % attempt Report creation
                reportFile = fullfile(resultsDir, 'Access_Report.txt');
                reportCmd = sprintf('ReportCreate */Chain/%s Type Save Style "Complete Chain Access" TimePeriod "%s" "%s" TimeStep 60 File "%s"', chainName, startTime, endTime, reportFile);
                try
                    root.ExecuteCommand(reportCmd);
                    % try convert
                    csvFile = fullfile(resultsDir, 'Access_Data.csv');
                    convertReportToCSV(reportFile, csvFile);
                    app.status.Text = '访问数据计算完成，CSV文件已保存';
                catch
                    % fallback: ComputeAccess
                    try
                        root.ExecuteCommand(sprintf('ComputeAccess */Chain/%s */Constellation/%s */Facility/%s', chainName, constName, createdStations{1}));
                        app.status.Text = '访问数据计算完成（ComputeAccess）';
                    catch
                        app.status.Text = '计算访问数据时遇到问题';
                    end
                end
            catch e
                warning('创建Chain或计算访问失败: %s', e.message);
            end
        end

        % optionally add cities
        if addCities
            try
                addChineseProvincialCapitals(root);
                app.status.Text = '正在添加中国省会城市...';
                drawnow;
            catch e
                warning('添加城市失败: %s', e.message);
            end
        end

        app.status.Text = sprintf('全部完成！成功导入 %d 颗卫星到STK场景 %s', satellitesCreated, scenName);
        app.status.FontColor = [0 0.4 0]; % 绿色成功
        updateProgress(1, '全部完成');
        drawnow;
    end

    %% ===== Utility functions =====

    function lines = parseTLEFile(tleFile)
        % returns N-by-3 cell array: {name, line1, line2}
        lines = {};
        fid = fopen(tleFile,'r');
        if fid == -1
            error('无法打开TLE文件: %s', tleFile);
        end
        raw = textscan(fid, '%s','Delimiter','\n','Whitespace','');
        fclose(fid);
        raw = raw{1};
        % remove empty lines
        raw = raw(~cellfun(@(s) isempty(strtrim(s)), raw));
        % Determine grouping
        if mod(length(raw),3) == 0
            n = length(raw)/3;
            lines = cell(n,3);
            for i=1:n
                idx = (i-1)*3 + 1;
                lines{i,1} = raw{idx};
                lines{i,2} = raw{idx+1};
                lines{i,3} = raw{idx+2};
            end
        elseif mod(length(raw),2) == 0
            n = length(raw)/2;
            lines = cell(n,3);
            for i=1:n
                idx = (i-1)*2 + 1;
                lines{i,1} = sprintf('Satellite_%d', i);
                lines{i,2} = raw{idx};
                lines{i,3} = raw{idx+1};
            end
        else
            idxs1 = find(cellfun(@(s) startsWith(strtrim(s), '1 '), raw));
            idxs2 = find(cellfun(@(s) startsWith(strtrim(s), '2 '), raw));
            n = min(length(idxs1), length(idxs2));
            lines = cell(n,3);
            for i=1:n
                l1 = raw{idxs1(i)};
                l2 = raw{idxs2(i)};
                lines{i,1} = sprintf('Satellite_%d', i);
                lines{i,2} = l1;
                lines{i,3} = l2;
            end
        end
    end

    function [uiap, root] = connectSTK(appLocal)
        % Try multiple COM ProgIDs. Save to app to keep references.
        if isfield(appLocal, 'STK_Handle') && ~isempty(appLocal.STK_Handle)
            uiap = appLocal.STK_Handle;
            root = appLocal.STK_Root;
            return;
        end

        stkVersions = {'STK13.application','STK12.application','STK11.application','STK.application'};
        uiap = [];
        root = [];
        lastErr = [];
        for v = 1:length(stkVersions)
            try
                uiap = actxserver(stkVersions{v});
                root = uiap.Personality2;
                appLocal.STK_Handle = uiap;
                appLocal.STK_Root = root;
                app.fig.UserData = appLocal; % persist
                break;
            catch ME
                lastErr = ME;
                continue;
            end
        end
        if isempty(root)
            error('无法连接任何STK版本: %s', lastErr.message);
        end
    end

    function addChineseProvincialCapitals(root)
        % Add China provincial capitals as Facilities
        caps = {
            'Beijing',39.9042,116.4074,0;
            'Tianjin',39.0842,117.20098,0;
            'Shijiazhuang',38.0428,114.5149,0;
            'Taiyuan',37.8706,112.5489,0;
            'Hohhot',40.8426,111.7492,0;
            'Shenyang',41.8057,123.4315,0;
            'Changchun',43.8171,125.3235,0;
            'Harbin',45.8038,126.5340,0;
            'Shanghai',31.2304,121.4737,0;
            'Nanjing',32.0603,118.7969,0;
            'Hangzhou',30.2741,120.1551,0;
            'Hefei',31.8206,117.2272,0;
            'Fuzhou',26.0745,119.2965,0;
            'Nanchang',28.6829,115.8582,0;
            'Jinan',36.6512,117.1201,0;
            'Zhengzhou',34.7466,113.6254,0;
            'Wuhan',30.5928,114.3055,0;
            'Changsha',28.2282,112.9388,0;
            'Guangzhou',23.1291,113.2644,0;
            'Nanning',22.8170,108.3669,0;
            'Haikou',20.0440,110.1999,0;
            'Chongqing',29.4316,106.9123,0;
            'Chengdu',30.5728,104.0668,0;
            'Guiyang',26.6470,106.6302,0;
            'Kunming',25.0389,102.7183,0;
            'Lhasa',29.6520,91.1721,0;
            'Xi_an',34.3416,108.9398,0;
            'Lanzhou',36.0611,103.8343,0;
            'Xining',36.6171,101.7782,0;
            'Yinchuan',38.4665,106.2587,0;
            'Urumqi',43.8256,87.6168,0;
            'HongKong',22.3193,114.1694,0;
            'Macao',22.1987,113.5439,0
            };
        sc = root.CurrentScenario;
        for i=1:size(caps,1)
            try
                cmd = sprintf('New / */Facility %s', caps{i,1});
                root.ExecuteCommand(cmd);
                cmd = sprintf('SetPosition */Facility/%s Geodetic %f %f UseTerrain', caps{i,1}, caps{i,2}, caps{i,3});
                root.ExecuteCommand(cmd);
            catch
                try
                    facObj = sc.Children.New('eFacility', caps{i,1});
                    facObj.Position.AssignGeodetic(caps{i,2}, caps{i,3}, caps{i,4});
                catch e
                    warning('无法在STK中添加城市 %s: %s', caps{i,1}, e.message);
                end
            end
        end
    end

    function convertReportToCSV(reportFile, csvFile)
        % Try to convert a simple whitespace-separated report to CSV
        try
            if ~exist(reportFile,'file')
                warning('报告文件不存在: %s', reportFile);
                return;
            end
            fidIn = fopen(reportFile,'r');
            if fidIn == -1
                warning('无法打开报告: %s', reportFile);
                return;
            end
            fidOut = fopen(csvFile,'w');
            if fidOut == -1
                fclose(fidIn);
                warning('无法创建CSV: %s', csvFile);
                return;
            end
            lineCount = 0;
            while ~feof(fidIn)
                line = fgetl(fidIn);
                lineCount = lineCount + 1;
                if isempty(line) || all(isspace(line))
                    continue;
                end
                % collapse whitespace
                line = regexprep(line, '\s+', ',');
                fprintf(fidOut, '%s\n', line);
            end
            fclose(fidIn);
            fclose(fidOut);
            fprintf('成功转换报告文件: %s -> %s (%d 行)\n', reportFile, csvFile, lineCount);
        catch err
            warning('转换报告失败: %s', err.message);
            if exist('fidIn','var') && fidIn>0, fclose(fidIn); end
            if exist('fidOut','var') && fidOut>0, fclose(fidOut); end
        end
    end


function encodedStr = custom_urlencode(str)
    % 自定义URL编码函数，替代Communications Toolbox中的urlencode
    % 输入: str - 要编码的字符串
    % 输出: encodedStr - URL编码后的字符串
    
    % 保留字符（不需要编码）
    unreservedChars = ['A':'Z', 'a':'z', '0':'9', '-', '_', '.', '~'];
    
    % 将字符串转换为ASCII码
    str = char(str);
    encodedStr = '';
    
    for i = 1:length(str)
        c = str(i);
        if any(c == unreservedChars)
            % 保留字符直接添加
            encodedStr = [encodedStr c]; %#ok<AGROW>
        else
            % 特殊字符进行编码
            hexVal = dec2hex(c, 2);
            encodedStr = [encodedStr '%' hexVal]; %#ok<AGROW>
        end
    end
end

end