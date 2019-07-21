function mape = MAPE(x, input, target,n)
disp(x);
 net = newrb(input,target,0,x',n,1);
 training = net(input);
 mape = mean((abs(target-training))./target);
 %nmse = mean((abs(target-training)).^2);
