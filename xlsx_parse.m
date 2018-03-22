function [timing, temp] = xlsx_parse(xlFile)

[~, sheets, ~] = xlsfinfo(xlFile);

timing = xlsread(xlFile, 'Data', 'A:A');
timing = timing/1000;

if ~isnan(find(ismember(sheets, 'Description'))) == 1
    [~, txt, ~] = xlsread(xlFile, 'Description');
    tMethod = find(ismember(txt, 'Probe'));
    if regexp(char(txt(tMethod(1),2)), 'CHEPS') == 1
        temp = xlsread(xlFile, 'Data', 'B:B');
    elseif regexp(char(txt(tMethod(1),2)), 'ATS') == 1
        temp = xlsread(xlFile, 'Data', 'C:C');
    end
end

if size(temp, 1) ~= size(timing, 1)
    n = abs(size(temp, 1) - size(timing, 1));
    if size(temp, 1) > size(timing, 1)
        temp = temp(n+1:end);
    elseif size(temp, 1) < size(timing, 1)
        timing = timing(n+1:end);
    end
end

end