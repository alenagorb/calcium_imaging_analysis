function out = extractCaData(directory, expduration, expconditions) 
% Function extractCaData extracts the changes in intensities over time 
% from excel files generated using 'Time Measurement' function in NIS
% Elements software and plots the results of the individual measurements
% over time.
% 
% The function takes 3 compulsory inputs:
% directory: specifies the working directory where the files to be analysed
% are stored e.g. directory = "C:\Users\~\*.xlsx".
% expduration: final time up to which the time measurements were taken in
% string format e.g. imaging for 5 mins means that expduration = '05:00.0'.
% expconditions: a cell array containing strings describing the
% experimental conditions/treatments used e.g. expcondition = 
% {'Control','Treatment_1','Treatment_2','Treatment_3'}.
% These will be used to label the
% final plots of the results.

% Error checking for required number of inputs
if nargin < 3
    error('Not enough input arguments.');
end

% Extract all .xlsx files from the given directory and store the info 
% on all .xlsx files as a struct
D = dir(directory);

% Store all the tables as cells in T selecting only 'Time_m_s_','ROIID','Intensity'
% columns
for i = 1:length(D)
    opts = detectImportOptions(D(i).name);
    opts.SelectedVariableNames = {'Time_m_s_','ROIID','Intensity'};
    T{i} = readtable(D(i).name, opts);
    T{i}.Time_m_s_ = datetime(T{i}.Time_m_s_,'ConvertFrom','datenum','Format','mm:ss.S');
%     Convert the time column from serial date number to desired time
%     format
end

% Remove rows with data after a timestamp specified by expduration 
for ii = 1:length(T)
    T{ii}(string(T{ii}.Time_m_s_) > expduration,:) = [];
end

% Convert 'ROIID' and 'Intensity' columns into arrays
for iii = 1:length(T)
    A{iii}=table2array(T{iii}(:,2:end));
end

% Stack the intensities from all neurites/somas in each dish side-by-side in arrays
Anew = cell(1,length(A));
npoints = length(A{1})/length(unique(A{1}(:,1)));
for n = 1:length(A)
    Tnew=zeros(npoints,max(unique(A{n}(:,1))));
    for a=1:max(unique(A{n}(:,1)))
        for b=1:height(A{n})
            if A{n}(b,1)== a
                Tnew(b-npoints*(a-1),a) = A{n}(b,2);
            end 
        end
    end
  Anew{n} = Tnew;  
end 

% Generate timestamps for all data points in Anew arrays
Time = T{1}.Time_m_s_(1:npoints); 

% Generate graphs for all neurites/somas for all conditions using
% expconditions as graph headings
graph_titles = expconditions;
for c = 1:length(Anew)
    subplot(floor(length(Anew)/2),3,c)
    plot(Time,Anew{c})
    title(graph_titles{c})
    xlabel('Time (mins)')
    ylabel('Intensity')
    xticks(Time(1:60*2:end))
    xticklabels(string(Time(1:60*2:end)))
end
