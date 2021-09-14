%% Batch File Processing for ContractQuantM2D %%

clear all;
close all;

%Set whether input files are nd2 (0) or Tif (1)

Tif = 0;

%Set whether the width of peri-event contractions should be adjusted from
%the default by the number of frames desired (a negative number will adjust
%for narrow contractions while a positive number will adjust for very long
%contractions
offsetMargin = 0;

% get filename info for images
% Select one file in the folder containing all files to be analyzed.
[filename, pathname, filterIndex] = uigetfile('*.*', 'Open Tissue/Cell Image');
cd(pathname);
direct=dir;

% create output directory
dirOut = 'ContractQuant_Results'; % Name of folder containing outputs.
mkdir(dirOut);

% establish excel file for outputs
experiment_title = 'ContractilityResults'; % Name of output excel file.
cd(dirOut);

headers = {'ImageNames','Tracking_Distance(um)','Pixel_Distance','Frequency','Max_FrShorten','Min_Length','Contract_Vel_20','Contract_Vel_50','Contract_Vel_80','Relax_Vel_80','Relax_Vel_50','Relax_Vel_20',...
    'NormContract_Vel_20','NormContract_Vel_50','NormContract_Vel_80','NormRelax_Vel_80','NormRelax_Vel_50','NormRelax_Vel_20','Max_Contract_Accel','Relax_Accel_Time','ContractTime',...
    'Contract_Decel_Time','Relax_Time','Event_Time','Max_Contract_Vel','Max_Relax_Vel','Velocity_Asymmetry_Index','StDev_Frequency','StDev_Max_FrShorten','StDev_Min_Length','StDev_Contract_Vel_20',...
    'StDev_Contract_Vel_50','StDev_Contract_Vel_80','StDev_Relax_Vel_80','StDev_Relax_Vel_50','StDev_Relax_Vel_20','StDev_Max Contraction_Acceleration','StDev_Relaxation_Acceleration',...
    'StDev_Contraction_Time','StDev_Contraction_Deceleration_Time','StDev_Relaxation_Time','StDev_Event_Time','StDev_Max_Contraction_Velocity','StDev_Max_Relaxation_Velocity'};

xlswrite(strcat(experiment_title,'.xls'),headers,1,'A1');

n=1; %count all files in directory
m=1; %count only analyzed files
for imgCount=1:length(direct)
    cd(pathname);
    if direct(imgCount).bytes < 10000   % Filter based on filesize
        n=n+1;                  % Skip small files (headers,text, & other)
    else                        % Continue if large enough to be an image
          File = direct(imgCount).name; %Get name of file being analyzed
      try  %the try and catch commands will skip contraction videos that trigger errors and keep processing subsequent files in the folder
         [TrackingDistance_um,PixelDistance,ContractionFrequency,median_max_FrShort,median_min_length,median_Contract_20Vel,median_Contract_50Vel,median_Contract_80Vel,median_Relax_80Vel,median_Relax_50Vel,median_Relax_20Vel,...
         median_Contract_20VelNorm,median_Contract_50VelNorm,median_Contract_80VelNorm,median_Relax_80VelNorm,median_Relax_50VelNorm,median_Relax_20VelNorm,median_max_ContractAccel,...
         median_Relaxation_Acceleration_time,median_Contraction_time,median_Contraction_Deceleration_time,median_Relaxation_time,median_event_time,median_max_Contract_Vel,median_max_Relax_Vel,...
         Velocity_Asymmetry_Index,std_inst_fre,std_max_FrShort,std_min_Length,std_Contract_20Vel,std_Contract_50Vel,std_Contract_80Vel,std_Relax_80Vel,std_Relax_50Vel,std_Relax_20Vel,std_max_ContractAccel,...
         std_Relaxation_Acceleration_time,std_Contraction_time,std_Contraction_Deceleration_time,std_Relaxation_time,std_event_time,std_max_Contract_Vel,std_max_Relax_Vel] = ContractQuantM2D(pathname,File,dirOut,Tif,offsetMargin);
        
        xlsOut = {File,TrackingDistance_um,PixelDistance,ContractionFrequency,median_max_FrShort,median_min_length,median_Contract_20Vel,median_Contract_50Vel,median_Contract_80Vel,median_Relax_80Vel,median_Relax_50Vel,median_Relax_20Vel,...
         median_Contract_20VelNorm,median_Contract_50VelNorm,median_Contract_80VelNorm,median_Relax_80VelNorm,median_Relax_50VelNorm,median_Relax_20VelNorm,median_max_ContractAccel,...
         median_Relaxation_Acceleration_time,median_Contraction_time,median_Contraction_Deceleration_time,median_Relaxation_time,median_event_time,median_max_Contract_Vel,median_max_Relax_Vel,...
         Velocity_Asymmetry_Index,std_inst_fre,std_max_FrShort,std_min_Length,std_Contract_20Vel,std_Contract_50Vel,std_Contract_80Vel,std_Relax_80Vel,std_Relax_50Vel,std_Relax_20Vel,std_max_ContractAccel,...
         std_Relaxation_Acceleration_time,std_Contraction_time,std_Contraction_Deceleration_time,std_Relaxation_time,std_event_time,std_max_Contract_Vel,std_max_Relax_Vel};
        
      catch ME
          fprintf('Contraction analysis not successful for %s - Error: %s\n',File,ME.message);
          ErrorMessage=sprintf('Contraction analysis not successful: %s. Inspect video for quality. May be able to re-run with adjustment of the tracking point coordinates.',ME.message);
          xlsOut = {File,ErrorMessage};
       end
      
       xlsRow = strcat('A',num2str(m+1));
       cd(dirOut);
       xlswrite(strcat(experiment_title,'.xls'),xlsOut,1,xlsRow);
       n=n+1;
       m=m+1;
    end
end
