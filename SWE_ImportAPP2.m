%GUI for adding functional data to 3DX-exported structure.

%Add filename created by Excel_Creation_3DX.m to lines 180, 292, 537 and
%571

clc;
clear;
%Clear command window.

Data = readtable('Barron''s Report.xlsx');
Height = height(Data);
%Import data from SWE export spreadsheet.

fig = uifigure;
fig.Name = 'Part Associator';
fig.Position = [100 100 1700 1000];
%New figure.

pnl = uipanel(fig,'Position',[1500 850 100 20]);
DataLabel = uilabel(fig,'Position',[1520 850 100 20],...
    'Text','');
%Panel for displaying selected cell.

T = uitable(fig,'Position',[20 20 1400 900],...
    'CellSelectionCallback',@(T,event) dataselect(T,event,fig,DataLabel));
%Plot Table.

T.Data = Data;
%Plot initial data to table.

figcaption = uilabel(fig,'Position',[20 920 1000 30]);
figcaption.Text = 'Select Component Data to Add';
figcaption.FontSize = 20;
%Instruction caption for selection table.

%Display selected data above select button.
function [Indices] = dataselect(T,event,fig,DataLabel)

Indices = event.Indices;
Selection = T.Data.Designator{Indices(1)};
DataLabel.Text = Selection;
%Assign selected cell data to display label.

uibutton('Parent',fig,...
    'Position',[1500 820 140 20],...
    'Text','Add to Existing',...
    'ButtonPushedFcn',@(but,event) confirmation(but,event,Indices,T,fig));
%Create button to confirm selection.

uibutton('Parent',fig,...
    'Position',[1500 780 140 20],...
    'Text','Create New',...
    'ButtonPushedFcn',@(but,event) confirmation2(but,event,Indices,T,fig));
%Create button to confirm selection.

end

%Use dialog box to confirm component selection.
function confirmation(~,~,Indices,T,fig)

selection = T.Data.Designator{Indices(1)};
selectDesc = T.Data.ComponentDescription{Indices(1)};
%Get component data to display on confirmation.

d = dialog('Position',[800 600 300 150],'Name','Confirm Selection');
%Dialog box to confirm selected component.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[45 90 210 40],...
    'String','Add this component data to a part?');
%Selection string.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[45 60 210 40],...
    'String',selection);
%Display selected Mark.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[45 40 210 40],...
    'String',selectDesc);
%Display selected Description.

uicontrol('Parent',d,...
    'Style','pushbutton',...
    'Position',[50 20 70 25],...
    'String','Yes',...
    'Callback',@(but,event) associate(but,event,Indices,T,fig));
%Confirm selection and open next figure.

uicontrol('Parent',d,...
    'Style','pushbutton',...
    'Position',[175 20 70 25],...
    'String','No',...
    'Callback','delete(gcf)');
%Return to part selection

end

%Use dialog box to confirm component selection.
function confirmation2(~,~,Indices,T,fig)

selection = T.Data.Designator{Indices(1)};
selectDesc = T.Data.ComponentDescription{Indices(1)};
%Get component data to display on confirmation.

d = dialog('Position',[800 600 300 150],'Name','Confirm Selection');
%Dialog box to confirm selected component.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[45 90 210 40],...
    'String','Create new 3D part data?');
%Selection string.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[45 60 210 40],...
    'String',selection);
%Display selected Mark.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[45 40 210 40],...
    'String',selectDesc);
%Display selected Description.

uicontrol('Parent',d,...
    'Style','pushbutton',...
    'Position',[50 20 70 25],...
    'String','Yes',...
    'Callback',@(but,event) associate2(but,event,Indices,T,fig));
%Confirm selection and open next figure.

uicontrol('Parent',d,...
    'Style','pushbutton',...
    'Position',[175 20 70 25],...
    'String','No',...
    'Callback','delete(gcf)');
%Return to part selection

end

%Save selected data to associate.
function [AssociateData] = associate(~,~,Indices,T,fig)

delete(gcf);
%Remove dialog box.

if strcmp(T.Data.Designator{Indices(1)},'Associated') == 1
    
    wrong = dialog('Position',[800 600 250 150],'Name','Error');
    
    uicontrol('Parent',wrong,'Style','text',...
        'Position',[20 100 210 40],...
        'String','Component already associated.');
    
    uicontrol('Parent',wrong,'Style','pushbutton',...
        'Position',[95 20 70 25],...
        'String','Continue',...
        'Callback','delete(gcf)');
    
else
    
    Mark = T.Data.Designator{Indices(1)};
    Description = T.Data.ComponentDescription{Indices(1)};
    Manufacturer = T.Data.Manufacturer{Indices(1)};
    PartNumber = T.Data.PartNumber{Indices(1)};
    %Get selected component data.

    AssociateData = {Mark,Description,Manufacturer,PartNumber};
    %Concatenate selected component data.

    T.Data.Designator{Indices(1)} = 'Associated';
    xlswrite('Barron''s Report.xlsx',T.Data.Designator,'4','B2');
    %Save new data to existing file.

    [~,importdata] = xlsread('D00019427 A.1_Reference.xlsx');
    %Read data from 3DX spreadsheet.

    newfig = uifigure;
    newfig.Name = 'Part Associator';
    newfig.Position = [100 100 1700 1000];
    %Create new figure.

    DataLabel1 = uilabel(newfig,'Position',[1450 620 200 20],...
        'Text','');
    DataLabel2 = uilabel(newfig,'Position',[1450 600 200 20],...
        'Text','');

    [~,sz2] = size(importdata);
    s = strings(sz2);
    cells = cellstr(s);

    Tab = uitable(newfig,'Position',[20 20 1400 900],...
        'CellSelectionCallback',@(Tab,event) displaypartdata(Tab,event,AssociateData,fig,newfig,DataLabel1,DataLabel2));
    Tab.Data = [importdata;cells];
    %Plot 3DX part data in new figure with cell select callback.

    caption = uilabel(newfig,'Position',[20 920 1000 30]);
    caption.Text = 'Select Part to Add Data Into';
    caption.FontSize = 20;
    %Instruction caption for new table.
    
end

end

%Save selected data to associate.
function associate2(~,~,Indices,T,fig)

delete(gcf);
%Remove dialog box.

if strcmp(T.Data.Designator{Indices(1)},'Associated') == 1
    
    wrong = dialog('Position',[800 600 250 150],'Name','Error');
    
    uicontrol('Parent',wrong,'Style','text',...
        'Position',[20 100 210 40],...
        'String','Component already associated.');
    
    uicontrol('Parent',wrong,'Style','pushbutton',...
        'Position',[95 20 70 25],...
        'String','Continue',...
        'Callback','delete(gcf)');
    
else
    
    d = dialog('Position',[800 600 300 150],'Name','Part Type');
    %Dialog box to confirm selected component.

    uicontrol('Parent',d,...
        'Style','text',...
        'Position',[45 90 210 40],...
        'String','Is this part a component? If not, click Continue');
    %Selection string.

    comp = uicontrol('Parent',d,...
        'Style','radiobutton',...
        'Position',[50 50 80 25],...
        'String','Component');
    %Confirm selection and open next figure.

    uicontrol('Parent',d,...
        'Style','pushbutton',...
        'Position',[112 20 70 25],...
        'String','Continue',...
        'Callback',@(but,event) typecreate(but,event,comp,Indices,T,fig));
    
end

end

%Create the Type field for new component addition.
function typecreate(~,~,comp,Indices,T,fig)

if comp.Value == 1
    
    typestring = 'Physical Product|3D Part';
    
elseif comp.Value == 0
    
    typestring = 'Physical Product';
    
end

delete(gcf);

Mark = T.Data.Designator{Indices(1)};
Description = T.Data.ComponentDescription{Indices(1)};
Manufacturer = upper(T.Data.Manufacturer{Indices(1)});
PartNumber = T.Data.PartNumber{Indices(1)};
%Get selected component data.

Title = strcat(Manufacturer,' -./.',upper(PartNumber));
Title = replace(Title,'./.',' ');
%Create Title field from part information.

Type = typestring;
%Create Type field for added part.

AssociateData = {Mark,Description,Manufacturer,PartNumber,Title,Type};
%Concatenate selected component data.

T.Data.Designator{Indices(1)} = 'Associated';
xlswrite('Barron''s Report.xlsx',T.Data.Designator,'4','B2');
%Save new data to existing file.

[~,importdata] = xlsread('D00019427 A.1_Reference.xlsx');
%Read data from 3DX spreadsheet.

newfig = uifigure;
newfig.Name = 'Part Associator';
newfig.Position = [100 100 1700 1000];
%Create new figure.

DataLabel1 = uilabel(newfig,'Position',[1450 620 200 20],...
    'Text','');
DataLabel2 = uilabel(newfig,'Position',[1450 600 200 20],...
    'Text','');

Tab = uitable(newfig,'Position',[20 20 1400 900],...
    'CellSelectionCallback',@(Tab,event) displaypartdata2(Tab,event,AssociateData,fig,newfig,DataLabel1,DataLabel2));
Tab.Data = importdata;
%Plot 3DX part data in new figure with cell select callback.

caption = uilabel(newfig,'Position',[20 920 1000 30]);
caption.Text = 'Select assembly to insert component';
caption.FontSize = 20;
%Instruction caption for new table.

end

%Display current selection in window.
function displaypartdata(Tab,event,AssociateData,fig,newfig,DataLabel1,DataLabel2)

Indices = event.Indices;
SelectionTitle = Tab.Data{Indices(1),7};
SelectionDescription = Tab.Data{Indices(1),4};

DataLabel1.Text = SelectionTitle;
DataLabel2.Text = SelectionDescription;

uibutton('Parent',newfig,...
    'Position',[1450 570 140 20],...
    'Text','Select Component',...
    'ButtonPushedFcn',@(but,event) finalize(Tab,event,AssociateData,fig,newfig,Indices));
%Create button to confirm selection.

end

%Display current selection in window.
function displaypartdata2(Tab,event,AssociateData,fig,newfig,DataLabel1,DataLabel2)

Indices = event.Indices;
SelectionTitle = Tab.Data{Indices(1),7};
SelectionDescription = Tab.Data{Indices(1),4};

DataLabel1.Text = SelectionTitle;
DataLabel2.Text = SelectionDescription;

uibutton('Parent',newfig,...
    'Position',[1450 570 140 20],...
    'Text','Select Component',...
    'ButtonPushedFcn',@(but,event) finalize2(Tab,event,AssociateData,fig,newfig,Indices));
%Create button to confirm selection.

end

%Open confirmation dialog.
function finalize(Tab,~,AssociateData,fig,newfig,Indices)

Ind = Indices;
%Get selected cell position.

if Ind(1) == 1
    
    wrong = dialog('Position',[800 600 250 150],'Name','Error');
    
    uicontrol('Parent',wrong,'Style','text',...
        'Position',[20 100 210 40],...
        'String','Invalid Data Selection');
    
    uicontrol('Parent',wrong,'Style','pushbutton',...
        'Position',[95 20 70 25],...
        'String','Continue',...
        'Callback','delete(gcf)');
    
elseif strcmp(Tab.Data{Ind(1),4},'3D Shape') == 1
    
    wrong = dialog('Position',[800 600 250 150],'Name','Error');
    
    uicontrol('Parent',wrong,'Style','text',...
        'Position',[20 100 210 40],...
        'String','Invalid Data Selection');
    
    uicontrol('Parent',wrong,'Style','pushbutton',...
        'Position',[95 20 70 25],...
        'String','Continue',...
        'Callback','delete(gcf)');
    
else

Title = Tab.Data{Ind(1),7};
Desc = Tab.Data{Ind(1),4};
%Get part data for display.

d = dialog('Position',[800 600 250 150],'Name','Confirm');
%Dialog box for part data confirmation.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[20 100 210 40],...
    'String','Add component data to this part?');
%Confirmation string.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[20 80 210 40],...
    'String',Title);
%Selected part title.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[20 60 210 40],...
    'String',Desc);
%Selected part description.

uicontrol('Parent',d,...
    'Style','pushbutton',...
    'Position',[50 20 70 25],...
    'String','Yes',...
    'Callback',@(b,event) inserter(b,event,AssociateData,Tab,Ind,fig,newfig));
%Confirm selected part and call function to add data.

uicontrol('Parent',d,...
    'Style','pushbutton',...
    'Position',[120 20 70 25],...
    'String','No',...
    'Callback','delete(gcf)');
%Return to part selection.

end

end

%Open confirmation dialog.
function finalize2(Tab,~,AssociateData,fig,newfig,Indices)

Ind = Indices;
%Get selected cell position.

if Ind(1) == 1
    
    wrong = dialog('Position',[800 600 250 150],'Name','Error');
    
    uicontrol('Parent',wrong,'Style','text',...
        'Position',[20 100 210 40],...
        'String','Invalid Data Selection');
    
    uicontrol('Parent',wrong,'Style','pushbutton',...
        'Position',[95 20 70 25],...
        'String','Continue',...
        'Callback','delete(gcf)');
    
elseif strcmp(Tab.Data{Ind(1),4},'3D Shape') == 1
    
    wrong = dialog('Position',[800 600 250 150],'Name','Error');
    
    uicontrol('Parent',wrong,'Style','text',...
        'Position',[20 100 210 40],...
        'String','Invalid Data Selection');
    
    uicontrol('Parent',wrong,'Style','pushbutton',...
        'Position',[95 20 70 25],...
        'String','Continue',...
        'Callback','delete(gcf)');
    
elseif strcmp(Tab.Data{Ind(1),4},'Physical Product|3D Part') == 1
    
    wrong = dialog('Position',[800 600 250 150],'Name','Error');
    
    uicontrol('Parent',wrong,'Style','text',...
        'Position',[20 100 210 40],...
        'String','Invalid Data Selection');
    
    uicontrol('Parent',wrong,'Style','pushbutton',...
        'Position',[95 20 70 25],...
        'String','Continue',...
        'Callback','delete(gcf)');
    
else

Title = Tab.Data{Ind(1),7};
Desc = Tab.Data{Ind(1),4};
%Get part data for display.

d = dialog('Position',[800 600 250 150],'Name','Confirm');
%Dialog box for part data confirmation.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[20 100 210 40],...
    'String','Insert part(s) under this assembly?');
%Confirmation string.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[20 80 210 40],...
    'String',Title);
%Selected part title.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[20 60 210 40],...
    'String',Desc);
%Selected part description.

uicontrol('Parent',d,...
    'Style','pushbutton',...
    'Position',[50 20 70 25],...
    'String','Yes',...
    'Callback',@(b,event) identifier(b,event,AssociateData,Tab,Ind,fig,newfig));
%Confirm selected part and call function to add data.

uicontrol('Parent',d,...
    'Style','pushbutton',...
    'Position',[120 20 70 25],...
    'String','No',...
    'Callback','delete(gcf)');
%Return to part selection.

end

end

%Add part data to spreadsheet.
function identifier(~,~,AssociateData,Tab,Ind,fig,newfig)

[H,~] = size(Tab.Data);

I = Tab.Data{Ind(1),1};
identity = strcat(I,'|1');
AD = [identity,AssociateData];
AD = char(AD);
Tab.Data{H+1,1} = AD(1,:);
Tab.Data{H+1,3} = AD(7,:);
Tab.Data{H+1,5} = AD(4,:);
Tab.Data{H+1,6} = AD(5,:);
Tab.Data{H+1,7} = AD(6,:);
Tab.Data{H+1,9} = AD(2,:);
Tab.Data{H+1,10} = AD(3,:);

xlswrite('D00019427 A.1_Reference.xlsx',Tab.Data,'Sheet1','A1');
%Save new data to existing file.

delete(gcf);
%Remove confirmation dialog box.

uibutton('Parent',newfig,...
    'Position',[1500 780 140 50],...
    'Text','Finish and Exit',...
    'ButtonPushedFcn',@(but,event) resaver(but,event,fig,newfig));
%Create button to close all figures.

uibutton('Parent',newfig,...
    'Position',[1500 700 140 50],...
    'Text','Return to Components',...
    'ButtonPushedFcn',@(b,event) componentselect(b,event,newfig));
%Create button to close part selector and return to component data.

end

%Store saved SWE data to selected 3DX part.
function inserter(~,~,input,Tab,Ind,fig,newfig)

AssociateData = char(input);
Tab.Data{Ind(1),9} = AssociateData(1,:);
Tab.Data{Ind(1),10} = AssociateData(2,:);
Tab.Data{Ind(1),5} = upper(AssociateData(3,:));
Tab.Data{Ind(1),6} = AssociateData(4,:);

title = strcat(upper(AssociateData(3,:)),' -./.',upper(AssociateData(4,:)));
title = replace(title,'./.',' ');
Tab.Data{Ind(1),7} = title;
%Save selected component data to selected part.

xlswrite('D00019427 A.1_Reference.xlsx',Tab.Data,'Sheet1','A1');
%Save new data to existing file.

delete(gcf);
%Remove confirmation dialog box.

uibutton('Parent',newfig,...
    'Position',[1500 780 140 50],...
    'Text','Finish and Exit',...
    'ButtonPushedFcn',@(but,event) resaver(but,event,fig,newfig));
%Create button to close all figures.

uibutton('Parent',newfig,...
    'Position',[1500 700 140 50],...
    'Text','Return to Components',...
    'ButtonPushedFcn',@(b,event) componentselect(b,event,newfig));
%Create button to close part selector and return to component data.

end

%Close all figures.
function resaver(~,~,fig,newfig)

delete(newfig);
delete(fig);

end

%Close part selection figure.
function componentselect(~,~,newfig)

delete(newfig);

end