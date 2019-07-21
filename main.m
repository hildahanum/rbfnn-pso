clear all
close all
clc
%% DATA INPUT
fileExcel = 'data.xlsx';
sheet = 'in';

xlRange = 'A1:A769';
dataTrainExcel = xlsread(fileExcel,sheet,xlRange)';
sizeTrain = size(dataTrainExcel);

xlRange = 'A770:A1025';
dataTestExcel = xlsread(fileExcel,sheet,xlRange)';



xlRange = 'B1:B1025';
dataValidationExcel = xlsread(fileExcel,sheet,xlRange)';

% data training
dataTrainInput = [dataTrainExcel(1:end-2); dataTrainExcel(2:end-1)];
dataTrainTarget = dataTrainExcel(3:end);

% data testing
dataTestInput = [dataTestExcel(1:end-2); dataTestExcel(2:end-1)];
dataTestTarget = dataTestExcel(3:end);

%data validation
dataValidationInput = [dataValidationExcel(1:end-2); dataValidationExcel(2:end-1)];
dataValidationTarget = dataValidationExcel(3:end);

%% TRAINING-TESTING
for n=2
    for it=[10]
        for pop=[5]

net = newrb(dataTrainInput,dataTrainTarget,0,13050,n,1);

simulasiNetTrainRBF = net(dataTrainInput);
MAPETrainingRBF = mean((abs(dataTrainTarget-simulasiNetTrainRBF))./dataTrainTarget);
MSETrainingRBF = mean((abs(dataTrainTarget-simulasiNetTrainRBF)).^2);

simulasiNetTestRBF = net(dataTestInput);
MAPETestingRBF = mean((abs(dataTestTarget-simulasiNetTestRBF))./dataTestTarget);
MSETestingRBF = mean((abs(dataTestTarget-simulasiNetTestRBF)).^2);


fungsiMAPE = @(x) MAPE(x, dataTrainInput, dataTrainTarget, n);
[x, err_ga] = pso(fungsiMAPE, it, pop);

net = newrb(dataTrainInput,dataTrainTarget,0,x',n,1);

simulasiNetTrainPSO = net(dataTrainInput);
MAPETrainingRBFPSO = mean((abs(dataTrainTarget-simulasiNetTrainPSO))./dataTrainTarget);
MSETrainingRBFPSO = mean((abs(dataTrainTarget-simulasiNetTrainPSO)).^2);

simulasiNetTestPSO = net(dataTestInput);
MAPETestingRBFPSO = mean((abs(dataTestTarget-simulasiNetTestPSO))./dataTestTarget);
MSETestingRBFPSO = mean((abs(dataTestTarget-simulasiNetTestPSO)).^2);


modeltestingrbf=['testingrbf','_',num2str(n),'_',num2str(it),'_',num2str(pop),'_','.mat'];
save(modeltestingrbf, 'simulasiNetTestRBF');

modeltrainingrbf=['trainingrbf','_',num2str(n),'_',num2str(it),'_',num2str(pop),'_','.mat'];
save(modeltrainingrbf, 'simulasiNetTrainRBF');

modeltrainingrbfpso=['trainingrbfpso','_',num2str(n),'_',num2str(it),'_',num2str(pop),'_','.mat'];
save(modeltrainingrbfpso, 'simulasiNetTrainPSO');

modeltestingrbfpso=['testingrbfpso','_',num2str(n),'_',num2str(it),'_',num2str(pop),'_','.mat'];
save(modeltestingrbfpso, 'simulasiNetTestPSO');

spread = x;

mapetrrbf = MAPETrainingRBF;
mapeterbf = MAPETestingRBF;
mapetrrbfpso = MAPETrainingRBFPSO;
mapeterbfpso = MAPETestingRBFPSO;

msetrrbf = MSETrainingRBF;
mseterbf = MSETestingRBF;
msetrrbfpso = MSETrainingRBFPSO;
mseterbfpso = MSETestingRBFPSO;

z=[num2str(n),'_',num2str(it),'_',num2str(pop)];
nilai = {z};

checkforfile=exist(strcat(pwd,'\','HasilOutput.csv'),'file');
if checkforfile==0;
    header1 = {'MAPE Training RBF'};
    header2 = {'MAPE Training RBF-PSO'};
    header3 = {'MAPE Testing RBF'};
    header4 = {'MAPE Testing RBF-PSO'};
    header5 = {'MSE Training RBF'};
    header6 = {'MSE Training RBF-PSO'};
    header7 = {'MSE Testing RBF'};
    header8 = {'MSE Testing RBF-PSO'};
    headerPerforma = {'Kode Model', 'Test MSE'};
    xlswrite('HasilOutput.csv', headerPerforma, 'HasilPerforma', 'A1');
    N = 0;
else
    N=size(xlsread('HasilOutput.csv', 'HasilPerforma'),1);
end
    xlswrite('HasilOutput.csv', spread,z,'A1');
    xlswrite('HasilOutput.csv', header1,z,'B1');
    xlswrite('HasilOutput.csv', mapetrrbf,z,'B2');
    xlswrite('HasilOutput.csv', header2,z,'C1');
    xlswrite('HasilOutput.csv', mapeterbf,z,'C2');
    xlswrite('HasilOutput.csv', header3,z,'D1');
    xlswrite('HasilOutput.csv', mapetrrbfpso,z,'D2');
    xlswrite('HasilOutput.csv', header4,z,'E1');
    xlswrite('HasilOutput.csv', mapeterbfpso,z,'E2');
    xlswrite('HasilOutput.csv', header5,z,'F1');
    xlswrite('HasilOutput.csv', msetrrbf,z,'F2');
    xlswrite('HasilOutput.csv', header6,z,'G1');
    xlswrite('HasilOutput.csv', mseterbf,z,'G2');
    xlswrite('HasilOutput.csv', header7,z,'H1');
    xlswrite('HasilOutput.csv', msetrrbfpso,z,'H2');
    xlswrite('HasilOutput.csv', header8,z,'I1');
    xlswrite('HasilOutput.csv', mseterbfpso,z,'I2');
  
    AA = strcat('A', num2str(N+2));
    xlswrite('HasilOutput.csv', nilai,'HasilPerforma',AA);
    
        end
    end
end

%% FORECAST

%data forecast
xlRange = 'A1024:A1240';
dataPrediksiExcel = xlsread(fileExcel,sheet,xlRange)';
dataPrediksiInput = [dataPrediksiExcel(1:end-1); dataPrediksiExcel(2:end)];
dataPrediksiTarget = dataPrediksiExcel(3:end);

for n=2
    sp= 26100;
%net = newrb(dataTrainInput,dataTrainTarget,0,sp,n,1);
simulasiNetPredict = net(dataPrediksiInput);
end

%% VALIDASI


simulasiNetValidation = net(dataValidationInput);

MAPEValidation = mean((abs(dataValidationTarget-simulasiNetValidation))./dataValidationTarget);
MSEValidation = mean((abs(dataValidationTarget-simulasiNetValidation)).^2);

modelvalidation=['validation','.mat'];
save(modelvalidation, 'simulasiNetValidation');

target=transpose(dataValidationTarget);
hasilSim=transpose(simulasiNetValidation);
z=num2str(n);
nilai={z};

checkforfile=exist(strcat(pwd,'\','Validation.csv'),'file');
if checkforfile==0
    header1 = {'Data Aktual'};
    header2 = {'Peramalan'};
    header3 = {'MAPE Validation'};
    header4 = {'MSE Validation'};
    headerPerforma = {'Kode Model', 'Test MSE'};
    xlswrite('Validation.csv', headerPerforma, 'HasilPerforma', 'A1');
    N = 0;
else
    N=size(xlsread('Validation.csv', 'HasilPerforma'),1);
end
    xlswrite('Validation.csv', header1,z,'A1');
    xlswrite('Validation.csv', target,z,'A2');
    xlswrite('Validation.csv', header2,z,'B1');
    xlswrite('Validation.csv', hasilSim,z,'B2');
    xlswrite('Validation.csv', header3,z,'C1');
    xlswrite('Validation.csv', MAPEValidation,z,'C2');
    xlswrite('Validation.csv', header4,z,'D1');
    xlswrite('Validation.csv', MSEValidation,z,'D2');

    AA = strcat('A', num2str(N+2));
    xlswrite('Validation.csv', nilai,'HasilPerforma',AA);
