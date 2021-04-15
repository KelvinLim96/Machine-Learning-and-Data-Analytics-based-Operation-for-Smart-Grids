%this code is to generate power system power flow data
%33 bus system = NTU campus
%DIP E015

clear all;
close all;

define_constants;

%read data
Read_Excel = readtable('NTU_Energy_Consumption.xlsx');
powerdata = table2array(Read_Excel(:,2:61));

%%
%load system
mpc = loadcase('case33bw_NTU_n');
BD = 2:14; %buildings are located from bus 2 to bus 14
PV = 17:19; %solar PVs are located from bus 17 to bus 32

nPV = size(PV',1);
nBD = size(BD',1);
nb = nPV + nBD; %total solar buses and building buses

non_slack = [BD PV]; %all buses except slack bus

[~, Yf, ~] = makeYbus(mpc); %Yfrom to compute the current

%extract power injection 
consumption = mpc.bus(BD,PD); %building consumption

%solve for solution
sol = runpf(mpc);
V = sol.bus(:,VM);

%store data
N = size(powerdata,1); %number of data points
data = zeros(N, 3*nb+1); %N rows for n data points; 
                         %3*nb for nb buses, each bus has 2 power values 
                         %and 1 voltage value;
                         %the last column records the system power loss

%generate date
for i = 1:length(powerdata)
    mpc.bus(BD,PD) = 0.01*(powerdata(i, 1:2:25))'; %BD_P
    mpc.bus(BD,QD) = 0.01*(powerdata(i, 2:2:26))'; %BD_Q

    mpc.bus(PV,PD) = 0.1*(-powerdata(i, 29:2:34))'; %PV_P
    mpc.bus(PV,QD) = (powerdata(i, 30:2:35))'; %PV_Q 
 
    sol_new = runpf(mpc); %run power flow
    V_new = sol_new.bus(non_slack,VM); %extract voltage solution
    
    %power losses
    Vsystem = sol_new.bus(:,VM).*exp(1i*sol_new.bus(:,VA)*180/pi);
    I = Yf * Vsystem;
    I_abs = abs(I);

    r_branch = mpc.branch(:,BR_R); 
    Ploss = (I_abs.^2)' * r_branch;
    
    %record the data
    data(i, 1:2:2*nb-1) = mpc.bus([BD, PV]', PD); %all P for BD+PV
    data(i, 2:2:2*nb) = mpc.bus([BD, PV]', QD); %all Q for BD+PV 
    data(i,2*nb+1:end-1) = V_new'; %all V for BD+PV
    data(i,end) = Ploss; %total power loss
end

%save data
save('data.mat','data')


%% Approximating voltage V-V0-C0 = M*P
XX = data(:, 1:2:2*nb-1);
V0 = ones(size(data(:,2*nb+1:end-1)))*1.1;
yy = data(:,2*nb+1:end-1)- V0; %V-V0

%construct M & C0 by looping
num_V  = size(yy,2);
M = zeros(num_V);
C0 = zeros(num_V,1);

 for i = 1:num_V
     YY = yy(:, i);
     mdl = fitlm(XX,YY);
     COEF = mdl.Coefficients.Estimate;
     M(:,i) = COEF(2:end);
     C0(i) = COEF(1);
 end

%% Approximating Ploss = L*(V-V0-C0)^2 + C
clear X;

new_C0 = [];
for i=1:size(data(:,2*nb+1:end-1),1)
    new_C0 = [new_C0;C0'];
end

V0 = mpc.bus(1,VM);
X = (data(:,2*nb+1:end-1)- V0 - new_C0).^2;
y = data(:,end);

mdl = fitlm(X,y);
Coef = mdl.Coefficients.Estimate;
L = Coef(2:end);
L = 1;
Beta0 = Coef(1);

%% Compute the optimum power building power loss
clear X;
P = data(:, 1:2:2*nb-1);
r_branch = mpc.branch(:,BR_R); 

PLC = zeros(288,1);
PLUC = zeros(288,1);

for j =1:288
    P_BD = P(j, 1:1:13);

    % Dividing M  = [M_BD M_C]
    M_BD = M(:,1:nBD);
    M_C = M(:,nBD+1:end);

    Q = M_BD*P_BD';
    H =  M_C'*M_C;
    R = 2*L*Q'*M_C;
    S = L*Q'*Q;
    X = -0.5*pinv(L*H)*R';
    X(3) = X(3)/2;

    % test optimal P_C
    mpcC = mpc;
    mpcC.bus(BD,PD) = 0.01*data(1, 1:2:2*nBD-1)'; %BD_P
    mpcC.bus(BD,QD) = 0.01*data(1, 2:2:2*nBD)'; %BD_Q

    mpcC.bus(PV,PD) = -0.1*X; %PV_P
    mpcC.bus(PV,QD) = 0;

    sol_C = runpf(mpcC); %run power flow
    V_C = sol_C.bus(non_slack,VM); %extract voltage solution

    %power losses
    VsystemC = sol_C.bus(:,VM).*exp(1i*sol_C.bus(:,VA)*180/pi);
    IC = Yf * VsystemC;
    IC_abs = abs(IC);
    
    PlossC = (IC_abs.^2)' * r_branch;
    PlossUC = data(j,end);
    
    PLC(j) = PlossC;
    PLUC(j) = PlossUC;
end

figure(1)
plot(1:288,PLUC,'red');
hold on
plot(1:288,PLC,'blue');

title('Power loss reduction Plot');
xlabel('iteration');
ylabel('Real power (w)');
legend('Uncontrolled power loss','Controlled power loss')

%% use GPR for learning voltages in terms of power injections

for i = 1:nb

x_observed = data(1:200, 1:nb*2);
y_observed = data(1:200, 2*nb+i); %NEC voltage

a = 'ardexponential';
 
V_NEC_model = fitrgp(x_observed,y_observed, 'KernelFunction', a);

T = data(201:288, 1:nb*2);

V_pre = predict(V_NEC_model,T);
error{i}  = max((abs(V_pre - data(201:end, 2*nb+2))./data(201:end, 2*nb+2)) * 100);

end

figure(2)
scatter(x_observed(:,1),y_observed,'r') 
hold on
y_pred = predict(V_NEC_model,x_observed);
scatter(x_observed(:,1),y_pred,'b');