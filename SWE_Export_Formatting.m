%Select Location BOM from SWE to import.

clc;
clear;
%Clear command window and variables.

Data = readtable('Barron''s Report.xlsx');
%Import data from SWE export spreadsheet.

Data = sortrows(Data,[5 2]);
%Sort parts by location.

Height = height(Data);
%Number of rows in SWE data.

Locations = Data.Location;
Locoptions = unique(Locations);
Loc = '+L1';
%Get Location options.

d = dialog('Position',[800 600 300 150],'Name','Location');
%Open location dialog box.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[45 90 210 40],...
    'String','Select schematic location to import.');
%Display string.

uicontrol('Parent',d,'Style','pushbutton',...
    'Position',[100 20 100 25],...
    'String','Continue',...
    'Callback','delete(gcf)');
%Close dialog.

uicontrol('Parent',d,'Style','popupmenu',...
    'Position',[100 50 100 25],...
    'String',Locoptions,...
    'Callback',@setloc);
%Set desired Location.

waitfor(d);
%Pause until figure is closed.

islocate = zeros(Height,1);
%Preallocate.

for k = 1:Height
        
    islocate(k) = strcmp(Data.Location{k},Loc);
    %Find data in selected location.
        
end

todelete = find(~islocate);
Data(todelete,:) = [];
%Select data only from desired location.

Height = height(Data);
%New table size.

Iden = cell(Height,1);
Name = cell(Height,1);
Display = cell(Height,1);
Des = Data.Designator;
PartDesc = Data.PartDescription;
Man = Data.Manufacturer;
Man = upper(Man);
PN = Data.PartNumber;
PartN = upper(PN);
CompDesc = Data.ComponentDescription;
Discipline = cell(Height,1);
Ext = cell(Height,1);
Inst = cell(Height,1);
%Define variables for import table.

pp = {'Physical Product|3D Part'};
Type = cell(Height,1);
Make = cell(Height,1);
%Preallocate.

for k = 1:Height
    
    Type(k,:) = pp;
    Iden(k,:) = {k};
    Make(k,:) = {'P'};
    
end
%Fill columns with the same value.

RootTitle = char(1);
title = strcat(Man,' -./.',PartN);
Title = replace(title,'./.',' ');
%Create reference title from part data.

d = dialog('Position',[800 600 300 150],'Name','Assembly');
%Open Title dialog box.

uicontrol('Parent',d,...
    'Style','text',...
    'Position',[45 90 210 40],...
    'String','Enter Assembly Title');
%Display string.

uicontrol('Parent',d,'Style','pushbutton',...
    'Position',[100 20 100 25],...
    'String','Continue',...
    'Callback','delete(gcf)');
%Close dialog.

uicontrol('Parent',d,'Style','edit',...
    'Position',[100 50 100 25],...
    'Callback',@setdisplay);
%Set Assembly Title.

waitfor(d);
%Pause until figure is closed.

Root = [{0},{''},{''},{'Physical Product'},{''},{''},{''},{''},{RootTitle}];
varNames = [{'Identifier'},{'(R) Name'},{'Display Name'},{'Type'},{'Extensions'},{'(R) EI_Discipline'},{'(R) EI_Manufacturer'},{'(R) EI_ManufacturerPN'},{'(R) Title'}];
T = [Iden,Name,Display,Type,Ext,Discipline,Man,PN,Title];
Export = [varNames;Root;T];
%Create table to write to file.

xlswrite('Importfile.xlsx',Export,'Sheet1','A1');
%Write data to new excel sheet.

%Assign desired location.
function setloc(source,~)

val = source.Value;
map = source.String;
Loc = map{val};
assignin('base','Loc',Loc);

end

%Assign top level assembly title.
function setdisplay(source,~)

val = source.String;
assignin('base','RootTitle',val);

end