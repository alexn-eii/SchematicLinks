%Create Reference Spreadsheet for association with Functional Reference
%Designators

%Input 3DX-export filename into readtable line (Line 11).
%Close Excel file before running anything else.

clc;
clear;
%Clear Workspace

importdata = readtable('D00019427-A.1_20200916_151139.xlsx');
%Import Excel Spreadsheet from 3DX

DisplayName = importdata.DisplayName;
TF = ismissing(DisplayName);
firstzero = find(TF,1);
n = firstzero - 1;
%Cut empty cells from imported data

DisplayName = importdata.DisplayName(1:n);
Identifier = importdata.Identifier(1:n);
Name = importdata.x_R_Name(1:n);
Path = importdata.Path(1:n);
Type = importdata.Type(1:n);
Extensions = importdata.Extensions(1:n);
Description = importdata.x_R_Description(1:n);
EI_Discipline = importdata.x_R_EI_Discipline(1:n);
EI_Manufacturer = importdata.x_R_EI_Manufacturer(1:n);
EI_ManufacturerPN = importdata.x_R_EI_ManufacturerPN(1:n);
Title = importdata.x_R_Title(1:n);
InstanceTitle = importdata.InstanceTitle(1:n);
%Split imported data into variables

Fullpath = strcat(Path,DisplayName);
%Create full path variable for RefDes

varNames = [{'Identifier'},{'Display Name'},{'Type'},{'(R) Description'},{'EI_Manufacturer'},{'EI_ManufacturerPN'},{'(R) Title'}];
%Create headers for export table

Typestring = strip(Type);
Typestring = char(Typestring);
%Get part type data for comparison.

todelete = zeros(n,1);
%Preallocate.

for k = 1:n
    
    if strcmp(strip(Typestring(k,:)),'Physical Product') == 1
        
        todelete(k,1) = 1;
        
    elseif strcmp(strip(Typestring(k,:)),'Physical Product|3D Part') == 1
        
        todelete(k,1) = 1;
        
    else
        
        todelete(k,1) = 0;
        
    end
    
end
%Find parts that do not need designators.

todelete = logical(~todelete);
DisplayName(todelete,:) = [];
Identifier(todelete,:) = [];
Name(todelete,:) = [];
Path(todelete,:) = [];
Type(todelete,:) = [];
Extensions(todelete,:) = [];
Description(todelete,:) = [];
EI_Discipline(todelete,:) = [];
EI_Manufacturer(todelete,:) = [];
EI_ManufacturerPN(todelete,:) = [];
Title(todelete,:) = [];
InstanceTitle(todelete,:) = [];
Fullpath(todelete,:) = [];
Typestring(todelete,:) = [];
%Remove parts without a Reference Designator.

T = [Identifier,DisplayName,Type,Description,EI_Manufacturer,EI_ManufacturerPN,Title];
[n,~] = size(T);
%Create new table structure.

singlelayer = strings([n,1]);
Refnumber = zeros(n,1);
FD = cellstr(strings([n,1]));
CD = cellstr(strings([n,1]));
%Preallocate variables.

for k = 1:n
    
    FD(k,:) = {''};
    CD(k,:) = {''};
    
    Refnumber (k,1) = k;
    
    if strcmp(strip(Typestring(k,:)),'Physical Product') == 1
        
        singlelayer(k,1) = '-AP';
        
    elseif strcmp(strip(Typestring(k,:)),'Physical Product|3D Part') == 1
        
        if strcmp(strip(EI_Discipline(k,:)),'Fluids') == 1
        
            singlelayer(k,1) = '-FC';
        
        elseif strcmp(strip(EI_Discipline(k,:)),'Electrical') == 1
            
            singlelayer(k,1) = '-EC';
            
        elseif strcmp(strip(EI_Discipline(k,:)),'Mechanical') == 1
            
            singlelayer(k,1) = '-MC';
            
        elseif strcmp(strip(EI_Discipline(k,:)),'Controls') == 1
            
            singlelayer(k,1) = '-CC';
            
        elseif strcmp(strip(EI_Discipline(k,:)),'Various') == 1
            
            singlelayer(k,1) = '-VC';
            
        else
            
            singlelayer(k,1) = '-C';
            
        end
        
    else
        
        singlelayer(k,1) = '';
        
    end
    
end
%Assign single layer reference designators to components.

Singlelayer = strcat(singlelayer,string(Refnumber));
%Assign single layer reference designators to components.

Refdesint = replace(Fullpath,DisplayName,Singlelayer);
ConcatRefDes = replace(Refdesint,'\-','.');
%Create concatenated product reference designators.

Main = [varNames;T];
%Create data table to write to excel.

ExtravarNames = [{'Concatenated Product Designator'},{'Function Designator'},{'Component Description'}];
Extravars = [ConcatRefDes,FD,CD];
Extra = [ExtravarNames;Extravars];
%Create functional variables.

Export = [Main,Extra];
%Combine tables for export.

filename = strcat(DisplayName(1,:),'_Reference.xlsx');
filename = char(filename);
%Create filename from top level assembly name.

xlswrite(filename,Export,'Sheet1','A1');
%Write data to new excel sheet

Excel = actxserver('Excel.Application');
%Call Excel COM object.

wkbk = Excel.Workbooks;
wdata = Open(wkbk,strcat('C:\Users\alexn\Documents\MATLAB\',filename));
wb = Excel.ActiveSheet;
%Create Worksheet objects.

Excel.Visible = 1;
%Open sheet.

wbRange = Range(wb,'A1');
wbRange.ColumnWidth = 20;
wbRange = Range(wb,'B1');
wbRange.ColumnWidth = 20;
wbRange = Range(wb,'C1');
wbRange.ColumnWidth = 25;
wbRange = Range(wb,'D1');
wbRange.ColumnWidth = 50;
wbRange = Range(wb,'E1');
wbRange.ColumnWidth = 20;
wbRange = Range(wb,'F1');
wbRange.ColumnWidth = 25;
wbRange = Range(wb,'G1');
wbRange.ColumnWidth = 35;
wbRange = Range(wb,'H1');
wbRange.ColumnWidth = 30;
wbRange = Range(wb,'I1');
wbRange.ColumnWidth = 20;
wbRange = Range(wb,'J1');
wbRange.ColumnWidth = 30;
cell = wb.Cells;
cell.NumberFormat = '@';
%Format Excel Sheet.

Save(wdata);
delete(Excel);
%Save and Quit Excel.
