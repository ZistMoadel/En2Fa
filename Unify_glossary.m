function Unify_glossary
clear
clc

% Thesis glossary ---------------------------------------------------------
[~,dat1] = xlsread('Unified_dic.xlsx'); % Based on existing theses
% Sorted_data = sortrows(dat1);
Eng_words = dat1(:,1);
Eng_words = strtrim(Eng_words);
Eng_words = lower(Eng_words);
Fa_words = dat1(:,2);
Fa_words = strtrim(Fa_words);
clear dat1
%Book glossary ------------------------------------------------------------
[RemoveAsk,Data] = xlsread('Glossary.xlsx'); % Based on other resources
NumelDiff = abs(numel(RemoveAsk) - numel(Data(:,1)));
if ~isempty(NumelDiff)
    RemoveAskT = zeros(size(Data,1),1);
    RemoveAskT(1:NumelDiff) = nan;
    RemoveAskT(NumelDiff+1:end) = RemoveAsk;
    RemoveAsk = RemoveAskT;
    clear RemoveAskT
end
RemoveIdx = find(RemoveAsk==1);
Eng = Data(:,1); Fa = Data(:,2);
Eng(RemoveIdx) = []; Fa(RemoveIdx) = []; % Remove redundant words from Book glossary: Identified by one "1" in the third column
Eng = strtrim(Eng);
Eng = lower(Eng);
Fa = strtrim(Fa);
clear Data RemoveAsk
% Concatenate two set of words
Eng_words = [Eng_words;Eng];
Fa_words = [Fa_words;Fa];
% Integration -------------------------------------------------------------
[unique_words,~,indx] = unique(Eng_words);
Repeated_words = unique_words(histc(indx,1:numel(indx))>1);

Rpt_Fa = ({}); RmvFromRptd = (0); ct1 = 1;
for ct = 1:numel(Repeated_words)
    Rpt_Temp = unique(Fa_words(ismember(Eng_words,Repeated_words{ct})));
    if numel(Rpt_Temp) == 1
        RmvFromRptd(ct1) = ct;
        ct1 = ct1 + 1;
    end
    Rpt_Fa{ct} = strjoin(Rpt_Temp','|');
    Fa_words(ismember(Eng_words,Repeated_words{ct})) = [];
    Eng_words(ismember(Eng_words,Repeated_words{ct})) = [];
    
end

%--------------------------------------------------------------------------
% Remove duplicates % NOTE: this lines assume that there are at most two
% duplicates for each redundant word 

% Duplicated_Engwords = Repeated_words(RmvFromRptd); % Remove words with one single meaning
% Duplicated_Fawords = Rpt_Fa(RmvFromRptd);
% 
% Find_duplicateFa = find(ismember(Fa_words,Duplicated_Fawords),numel(RmvFromRptd),'first');
% Find_duplicateEn = find(ismember(Eng_words,Duplicated_Engwords),numel(RmvFromRptd),'first');
% Fa_words(Find_duplicateFa) = []; % Keep only one word among duplicates
% Eng_words(Find_duplicateEn) = []; 
% Repeated_words(RmvFromRptd) = [];
% Rpt_Fa(RmvFromRptd) = [];
%--------------------------------------------------------------------------
New_len = numel(Eng_words)+1:numel(Eng_words)+numel(Rpt_Fa);
Eng_words(New_len) = Repeated_words;
Repeated_words = regexprep(Repeated_words,'^\s*.','${upper($0)}'); % Capitalize first letter
Eng_words = regexprep(Eng_words,'^\s*.','${upper($0)}'); % Capitalize first letter
[Eng_words,sortIdx] = sort(Eng_words); % sort Alphabetically 
[~,Where_repeated_words] = intersect(Eng_words,Repeated_words);

Fa_words(New_len) = Rpt_Fa;
Fa_words = Fa_words(sortIdx);
Show_changed = nan(numel(Fa_words),1);
Show_changed(Where_repeated_words) = 1;

% Final check to remove duplicates
[unique_words,~,indx] = unique(Eng_words);
Repeated_words1 = unique_words(histc(indx,1:numel(indx))>1);

% Mix two glossaries ------------------------------------------------------
xlswrite('Unified_Glossary.xlsx',Eng_words,1,'A1')
xlswrite('Unified_Glossary.xlsx',Fa_words,1,'B1')
xlswrite('Unified_Glossary.xlsx',Show_changed,1,'C1')


