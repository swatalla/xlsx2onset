
%% xlsx2onset script for creating Onsets(n).mat file
% Written by Sebastian Atalla
% TODO: Java/Python interop for reading excel files
% - potential OS agnostic resolution
% xlsread() does not work on Mac due to ActiveX/COM dependence
% Res: Maven pom.xml to use Apache POI to read xlsx files
% Res: Explore Mono/.NET Core API for Excel interop
% Res: Pandas/openpyxl/xlrd w/ Python; need to be able to conv. data types

clear

xlDir = fullfile('X:\Research_Data\KL2_Subject_Data\Timing_Files\IASP_TIMING');
subjs = cell2mat(inputdlg('Enter 1 subject per line', 'Subjects', 20,{'' ''}, 'on'));

for t = 1:size(subjs, 1)
    xlSubj = subjs(t, 1:end);
    xlRoot = dir(fullfile(xlDir, xlSubj, '*RUN*.xls*'));
    
    for k = 1:length(xlRoot)
        
        clf
        %clearvars -except subjs xlDir xlSubj xlRoot k t
        xlFile = fullfile(xlDir, xlSubj, xlRoot(k).name);
        
        [~, xlName, ~] = fileparts(xlFile);
        
        if exist(fullfile(xlDir,xlSubj, 'Onsets'), 'dir') == 0
            mkdir(fullfile(xlDir,xlSubj, 'Onsets'))
        end
        
        ons_name = sprintf('%s',xlName(10:13),'Onsets',xlName(end),'.mat');
        
        if exist(fullfile(xlDir, xlSubj, 'Onsets', ons_name), 'file')
            ons_name = sprintf('%s',xlName(10:13),'Onsets',xlName(end),'_COPY','.mat');
        end
        
        names = {'warm', 'mild', 'mod', 'ramps'};
        onsets = {};
        durations = {};
        
        [timing, temp] = xlsx_parse(xlFile);
        
        truncTemp = medfilt1(temp, round(length(temp)/25), 'truncate');
        sgTemp = round(smoothdata(truncTemp, 'sgolay', 35));
        
        [n, bin] = hist(sgTemp, unique(sgTemp));
        [~, idx] = sort(-n);
        tBin = bin(idx);
        tStim = sort([tBin(1), tBin(2), tBin(3), tBin(4)]);
        
        sgTemp(sgTemp < tStim(1)) = tStim(1);
        sgTemp(sgTemp > tStim(4)) = tStim(4);
        
        %% First Block
        
        onset1_start = onset_start(1, sgTemp, tStim);
        
        if timing(onset1_start) > 24.1
            z_iter = 1;
            while timing(onset1_start) - timing(z_iter) > 24
                z_iter = z_iter + 1;
            end
            tOffset = timing(z_iter) - timing(1);
            temp = temp(z_iter:end);
            sgTemp = sgTemp(z_iter:end);
            timing = timing(z_iter:end);
            timing = timing - tOffset;
            
            onset1_start = onset_start(1, sgTemp, tStim);
            
        elseif timing(onset1_start) < 24
            tOffset = (diff([timing(onset1_start), 24])-timing(1));
            timing = timing + tOffset;
            
        else
            
        end
        
        onset1_stop = onset_end(onset1_start, sgTemp, tStim);
        onset1_duration = diff([timing(onset1_start), timing(onset1_stop)]);
        
        offset1_start = offset_start(onset1_stop, sgTemp, tStim);
        offset1_stop = offset_end(offset1_start, sgTemp, tStim);
        offset1_duration = diff([timing(offset1_start), timing(offset1_stop)]);
        
        baseline1_duration = timing(onset1_start);
        stim1_duration = timing(offset1_start) - timing(onset1_stop);
        
        %% Second Block
        
        onset2_start = onset_start(offset1_stop, sgTemp, tStim);
        onset2_stop = onset_end(onset2_start, sgTemp, tStim);
        onset2_duration = diff([timing(onset2_start), timing(onset2_stop)]);
        
        offset2_start = offset_start(onset2_stop, sgTemp, tStim);
        offset2_stop = offset_end(offset2_start, sgTemp, tStim);
        offset2_duration = diff([timing(offset2_start), timing(offset2_stop)]);
        
        baseline2_duration = timing(onset2_start) - timing(offset1_stop);
        stim2_duration = timing(offset2_start) - timing(onset2_stop);
        
        %% Third Block
        
        onset3_start = onset_start(offset2_stop, sgTemp, tStim);
        onset3_stop = onset_end(onset3_start, sgTemp, tStim);
        onset3_duration = diff([timing(onset3_start), timing(onset3_stop)]);
        
        offset3_start = offset_start(onset3_stop, sgTemp, tStim);
        offset3_stop = offset_end(offset3_start, sgTemp, tStim);
        offset3_duration = diff([timing(offset3_start), timing(offset3_stop)]);
        
        baseline3_duration = timing(onset3_start) - timing(offset2_stop);
        stim3_duration = timing(offset3_start) - timing(onset3_stop);
        
        %% Fourth Block
        
        onset4_start = onset_start(offset3_stop, sgTemp, tStim);
        onset4_stop = onset_end(onset4_start, sgTemp, tStim);
        onset4_duration = diff([timing(onset4_start), timing(onset4_stop)]);
        
        offset4_start = offset_start(onset4_stop, sgTemp, tStim);
        offset4_stop = offset_end(offset4_start, sgTemp, tStim);
        offset4_duration = diff([timing(offset4_start), timing(offset4_stop)]);
        
        baseline4_duration = timing(onset4_start) - timing(offset3_stop);
        stim4_duration = timing(offset4_start) - timing(onset4_stop);
        
        %% Fifth Block
        
        onset5_start = onset_start(offset4_stop, sgTemp, tStim);
        onset5_stop = onset_end(onset5_start, sgTemp, tStim);
        onset5_duration = diff([timing(onset5_start), timing(onset5_stop)]);
        
        offset5_start = offset_start(onset5_stop, sgTemp, tStim);
        offset5_stop = offset_end(offset5_start, sgTemp, tStim);
        offset5_duration = diff([timing(offset5_start), timing(offset5_stop)]);
        
        baseline5_duration = timing(onset5_start) - timing(offset4_stop);
        stim5_duration = timing(offset5_start) - timing(onset5_stop);
        
        %% Sixth Block
        
        onset6_start = onset_start(offset5_stop, sgTemp, tStim);
        onset6_stop = onset_end(onset6_start, sgTemp, tStim);
        onset6_duration = diff([timing(onset6_start), timing(onset6_stop)]);
        
        offset6_start = offset_start(onset6_stop, sgTemp, tStim);
        offset6_stop = offset_end(offset6_start, sgTemp, tStim);
        offset6_duration = diff([timing(offset6_start), timing(offset6_stop)]);
        
        baseline6_duration = timing(onset6_start) - timing(offset5_stop);
        stim6_duration = timing(offset6_start) - timing(onset6_stop);
        
        %% Write Onsets
        
        if regexp(xlName(end-4:end), '_RUN1') == 1
            ml1 = [timing(onset1_stop), timing(onset1_start), timing(offset1_start), timing(offset1_stop), ...
                stim1_duration, baseline1_duration, onset1_duration, offset1_duration];
            md1 = [timing(onset2_stop), timing(onset2_start), timing(offset2_start), timing(offset2_stop), ...
                stim2_duration, baseline2_duration, onset2_duration, offset2_duration];
            wm1 = [timing(onset3_stop), timing(onset3_start), timing(offset3_start), timing(offset3_stop), ...
                stim3_duration, baseline3_duration, onset3_duration, offset3_duration];
            md2 = [timing(onset4_stop), timing(onset4_start), timing(offset4_start), timing(offset4_stop), ...
                stim4_duration, baseline4_duration, onset4_duration, offset4_duration];
            wm2 = [timing(onset5_stop), timing(onset5_start), timing(offset5_start), timing(offset5_stop), ...
                stim5_duration, baseline5_duration, onset5_duration, offset5_duration];
            ml2 = [timing(onset6_stop), timing(onset6_start), timing(offset6_start), timing(offset6_stop), ...
                stim6_duration, baseline6_duration, onset6_duration, offset6_duration];
            
        elseif regexp(xlName(end-4:end), '_RUN2') == 1
            md1 = [timing(onset1_stop), timing(onset1_start), timing(offset1_start), timing(offset1_stop), ...
                stim1_duration, baseline1_duration, onset1_duration, offset1_duration];
            wm1 = [timing(onset2_stop), timing(onset2_start), timing(offset2_start), timing(offset2_stop), ...
                stim2_duration, baseline2_duration, onset2_duration, offset2_duration];
            ml1 = [timing(onset3_stop), timing(onset3_start), timing(offset3_start), timing(offset3_stop), ...
                stim3_duration, baseline3_duration, onset3_duration, offset3_duration];
            wm2 = [timing(onset4_stop), timing(onset4_start), timing(offset4_start), timing(offset4_stop), ...
                stim4_duration, baseline4_duration, onset4_duration, offset4_duration];
            ml2 = [timing(onset5_stop), timing(onset5_start), timing(offset5_start), timing(offset5_stop), ...
                stim5_duration, baseline5_duration, onset5_duration, offset5_duration];
            md2 = [timing(onset6_stop), timing(onset6_start), timing(offset6_start), timing(offset6_stop), ...
                stim6_duration, baseline6_duration, onset6_duration, offset6_duration];
            
        elseif regexp(xlName(end-4:end), '_RUN3') == 1
            md1 = [timing(onset1_stop), timing(onset1_start), timing(offset1_start), timing(offset1_stop), ...
                stim1_duration, baseline1_duration, onset1_duration, offset1_duration];
            wm1 = [timing(onset2_stop), timing(onset2_start), timing(offset2_start), timing(offset2_stop), ...
                stim2_duration, baseline2_duration, onset2_duration, offset2_duration];
            ml1 = [timing(onset3_stop), timing(onset3_start), timing(offset3_start), timing(offset3_stop), ...
                stim3_duration, baseline3_duration, onset3_duration, offset3_duration];
            wm2 = [timing(onset4_stop), timing(onset4_start), timing(offset4_start), timing(offset4_stop), ...
                stim4_duration, baseline4_duration, onset4_duration, offset4_duration];
            ml2 = [timing(onset5_stop), timing(onset5_start), timing(offset5_start), timing(offset5_stop), ...
                stim5_duration, baseline5_duration, onset5_duration, offset5_duration];
            md2 = [timing(onset6_stop), timing(onset6_start), timing(offset6_start), timing(offset6_stop), ...
                stim6_duration, baseline6_duration, onset6_duration, offset6_duration];
            
        elseif regexp(xlName(end-4:end), '_RUN4') == 1
            md1 = [timing(onset1_stop), timing(onset1_start), timing(offset1_start), timing(offset1_stop), ...
                stim1_duration, baseline1_duration, onset1_duration, offset1_duration];
            ml1 = [timing(onset2_stop), timing(onset2_start), timing(offset2_start), timing(offset2_stop), ...
                stim2_duration, baseline2_duration, onset2_duration, offset2_duration];
            wm1 = [timing(onset3_stop), timing(onset3_start), timing(offset3_start), timing(offset3_stop), ...
                stim3_duration, baseline3_duration, onset3_duration, offset3_duration];
            ml2 = [timing(onset4_stop), timing(onset4_start), timing(offset4_start), timing(offset4_stop), ...
                stim4_duration, baseline4_duration, onset4_duration, offset4_duration];
            md2 = [timing(onset5_stop), timing(onset5_start), timing(offset5_start), timing(offset5_stop), ...
                stim5_duration, baseline5_duration, onset5_duration, offset5_duration];
            wm2 = [timing(onset6_stop), timing(onset6_start), timing(offset6_start), timing(offset6_stop), ...
                stim6_duration, baseline6_duration, onset6_duration, offset6_duration];
            
        end
        
        onsets{1} = [wm1(1); wm2(1)];
        onsets{2} = [ml1(1); ml2(1)];
        onsets{3} = [md1(1); md2(1)];
        onsets{4} = [wm1(2); wm2(2); wm1(3); wm2(3); ...
            ml1(2); ml2(2); ml1(3); ml2(3); ...
            md1(2); md2(2); md1(3); md2(3)];
        
        durations{1} = [wm1(5); wm2(5)];
        durations{2} = [ml1(5); ml2(5)];
        durations{3} = [md1(5); md2(5)];
        durations{4} = [wm1(7); wm2(7); wm1(8); wm2(8); ...
            ml1(7); ml2(7); ml1(8); ml2(8); ...
            md1(7); md2(7); md1(8); md2(8)];
        
        save(fullfile(xlDir, xlSubj, 'Onsets', ons_name), 'names', 'onsets', 'durations')
        
        %% Plot
        plot(timing, temp, timing, sgTemp)
        title(xlName, 'Interpreter', 'none')
        hold on
        line([timing(onset1_start), timing(onset1_start)],[temp(onset1_start)-10, temp(onset1_start)+10])
        line([timing(onset1_stop), timing(onset1_stop)],[temp(onset1_stop)-10, temp(onset1_stop)+10])
        line([timing(offset1_start), timing(offset1_start)],[temp(offset1_start)-10, temp(offset1_start)+10])
        line([timing(offset1_stop), timing(offset1_stop)],[temp(offset1_stop)-10, temp(offset1_stop)+1])
        line([timing(onset2_start), timing(onset2_start)],[temp(onset2_start)-10, temp(onset2_start)+10])
        line([timing(onset2_stop), timing(onset2_stop)],[temp(onset2_stop)-10, temp(onset2_stop)+10])
        line([timing(offset2_start), timing(offset2_start)],[temp(offset2_start)-10, temp(offset2_start)+10])
        line([timing(offset2_stop), timing(offset2_stop)],[temp(offset2_stop)-10, temp(offset2_stop)+10])
        line([timing(onset3_start), timing(onset3_start)],[temp(onset3_start)-10, temp(onset3_start)+10])
        line([timing(onset3_stop), timing(onset3_stop)],[temp(onset3_stop)-10, temp(onset3_stop)+10])
        line([timing(offset3_start), timing(offset3_start)],[temp(offset3_start)-10, temp(offset3_start)+10])
        line([timing(offset3_stop), timing(offset3_stop)],[temp(offset3_stop)-10, temp(offset3_stop)+10])
        line([timing(onset4_start), timing(onset4_start)],[temp(onset4_start)-10, temp(onset4_start)+10])
        line([timing(onset4_stop), timing(onset4_stop)],[temp(onset4_stop)-10, temp(onset4_stop)+10])
        line([timing(offset4_start), timing(offset4_start)],[temp(offset4_start)-10, temp(offset4_start)+10])
        line([timing(offset4_stop), timing(offset4_stop)],[temp(offset4_stop)-10, temp(offset4_stop)+10])
        line([timing(onset5_start), timing(onset5_start)],[temp(onset5_start)-10, temp(onset5_start)+10])
        line([timing(onset5_stop), timing(onset5_stop)],[temp(onset5_stop)-10, temp(onset5_stop)+10])
        line([timing(offset5_start), timing(offset5_start)],[temp(offset5_start)-10, temp(offset5_start)+10])
        line([timing(offset5_stop), timing(offset5_stop)],[temp(offset5_stop)-10, temp(offset5_stop)+10])
        line([timing(onset6_start), timing(onset6_start)],[temp(onset6_start)-10, temp(onset6_start)+10])
        line([timing(onset6_stop), timing(onset6_stop)],[temp(onset6_stop)-10, temp(onset6_stop)+10])
        line([timing(offset6_start), timing(offset6_start)],[temp(offset6_start)-10, temp(offset6_start)+10])
        line([timing(offset6_stop), timing(offset6_stop)],[temp(offset6_stop)-10, temp(offset6_stop)+10])
        plot(timing(onset1_start:onset1_stop), temp(onset1_start:onset1_stop), 'r')
        plot(timing(offset1_start:offset1_stop), temp(offset1_start:offset1_stop), 'r')
        plot(timing(onset2_start:onset2_stop), temp(onset2_start:onset2_stop), 'r')
        plot(timing(offset2_start:offset2_stop), temp(offset2_start:offset2_stop), 'r')
        plot(timing(onset3_start:onset3_stop), temp(onset3_start:onset3_stop), 'r')
        plot(timing(offset3_start:offset3_stop), temp(offset3_start:offset3_stop), 'r')
        plot(timing(onset4_start:onset4_stop), temp(onset4_start:onset4_stop), 'r')
        plot(timing(offset4_start:offset4_stop), temp(offset4_start:offset4_stop), 'r')
        plot(timing(onset5_start:onset5_stop), temp(onset5_start:onset5_stop), 'r')
        plot(timing(offset5_start:offset5_stop), temp(offset5_start:offset5_stop), 'r')
        plot(timing(onset6_start:onset6_stop), temp(onset6_start:onset6_stop), 'r')
        plot(timing(offset6_start:offset6_stop), temp(offset6_start:offset6_stop), 'r')
        plot(timing(onset1_stop:offset1_start), temp(onset1_stop:offset1_start), 'g')
        plot(timing(onset2_stop:offset2_start), temp(onset2_stop:offset2_start), 'g')
        plot(timing(onset3_stop:offset3_start), temp(onset3_stop:offset3_start), 'g')
        plot(timing(onset4_stop:offset4_start), temp(onset4_stop:offset4_start), 'g')
        plot(timing(onset5_stop:offset5_start), temp(onset5_stop:offset5_start), 'g')
        plot(timing(onset6_stop:offset6_start), temp(onset6_stop:offset6_start), 'g')
        
        saveas(gcf, fullfile(xlDir, xlSubj, sprintf('%s', xlName, '.png')))
        
    end
end