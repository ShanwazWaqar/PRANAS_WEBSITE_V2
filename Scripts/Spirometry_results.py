import warnings
warnings.filterwarnings('ignore')
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import find_peaks
import mpld3
import ast
import io
import base64
#from scipy.integrate import cumtrapz
from scipy.integrate import cumulative_trapezoid




def fvc_calc(data):
    # Get utility data
    imvmin, mvmin, imvmax, mvmax, mvAvgPow, mvAvgX, mvAvgY, master_data = utli_fun(data)
    
    # Check if we have enough minimum points in the data
    # We need at least 5 points if we're using index 3 and 4 (i and i+1)
    if len(imvmin) < 5:
        # If not enough points, use the first two peaks if available
        if len(imvmin) >= 2:
            i = 0  # Use the first point
        else:
            # Create an error plot if we don't have enough data
            plt.figure()
            plt.text(0.5, 0.5, f"Error: Not enough breath cycles detected in data.\nFound {len(imvmin)} valleys, need at least 5.", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=plt.gca().transAxes, fontsize=14, color='red')
            plt.xlabel('Time (s)')
            plt.ylabel('Pow')
            plt.grid(True)
            
            img_bytes_io_1 = io.BytesIO()
            plt.savefig(img_bytes_io_1)
            img_bytes_io_1.seek(0)
            img_base64_1 = base64.b64encode(img_bytes_io_1.read()).decode('utf-8')
            plt.close()
            
            # Return the error plot for both images
            return img_base64_1, img_base64_1
    else:
        i = 3  # Use the original index if we have enough data
    
    # First plot - show the breath pattern with marked points
    plt.figure()
    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmax / 1000, mvmax, 'or')
    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmin / 1000, mvmin, 'ob')
    
    # Safe indexing
    imin = imvmin[i]
    imin2 = imvmin[i+1]
    
    # Plot markers for exhalation start and end
    plt.plot(imvmin[i]/1000, mvmin[imin], '|', markersize=50, linewidth=50)
    plt.text(imvmin[i]/1000, mvmin[imin], 'Exhalation Start', fontsize=10, color='red', va='bottom', ha='right')
    plt.plot(imvmin[i+1]/1000, mvmin[imin2], '|', markersize=50, linewidth=50)
    plt.text(imvmin[i+1]/1000, mvmin[imin2], 'Inhalation End', fontsize=10, color='red', va='bottom', ha='right')

    plt.xlabel('Time (s)')
    plt.ylabel('Pow')
    plt.grid(True)
    plt.minorticks_on()
    
    # Save the first plot
    img_bytes_io_1 = io.BytesIO()
    plt.savefig(img_bytes_io_1)
    img_bytes_io_1.seek(0)
    img_base64_1 = base64.b64encode(img_bytes_io_1.read()).decode('utf-8')
    plt.close()
    
    # Second plot - flow volume curve
    try:
        plt.figure()
        time = imvmin[i+1] - imvmin[i]
        
        # Make sure we don't try to access values beyond array size
        if time <= 0 or imvmin[i] + time > len(mvAvgPow):
            raise ValueError("Invalid time range for FVC calculation")
            
        # Get time samples for the breath cycle
        tsVar = master_data["Samples"].iloc[0:time] / 1000
        
        # Get flow for the breath cycle
        if imvmin[i] >= len(mvAvgPow) or imvmin[i+1] > len(mvAvgPow):
            raise ValueError("Index out of bounds in breath cycle data")
            
        flow = mvAvgPow[imvmin[i]:imvmin[i+1]]
        
        # Ensure we have enough samples in tsVar and flow
        if len(tsVar) < 2 or len(flow) < 2:
            raise ValueError("Not enough samples for flow calculation")
            
        # Ensure tsVar and flow have the same length
        min_length = min(len(tsVar), len(flow))
        tsVar = tsVar[:min_length]
        flow = flow[:min_length]
        
        # Calculate flow rate
        flow_rate = np.diff(flow) / np.diff(tsVar)
        flow_df = pd.DataFrame({
            'flow_rate': flow_rate
        })
        
        # Check for sufficient data for trapezoid integration
        if len(flow_rate) < 2 or len(tsVar) < 3:
            raise ValueError("Not enough data points for volume calculation")
            
        # Calculate flow volume using trapezoidal integration
        flow_volume = cumulative_trapezoid(flow_rate) * np.diff(tsVar[:-1])
        
        # Check if we got a valid flow volume
        if len(flow_volume) == 0:
            raise ValueError("Empty flow volume array")
            
        fvc = max(flow_volume)

        # Smoothing flow rate using moving average
        # Use a smaller window if we don't have enough points
        window_size = min(750, len(flow_rate) // 2)
        if window_size < 1:
            window_size = 1
            
        flow_rate_smoothed = flow_df['flow_rate'].rolling(window=window_size, min_periods=1).mean()
        flow_rate_smoothed = flow_rate_smoothed.values
        
        # Ensure smoothed flow rate and flow volume have compatible lengths
        min_len = min(len(flow_volume), len(flow_rate_smoothed) - 1)
        
        # Plotting
        plt.plot(flow_volume[:min_len], flow_rate_smoothed[:min_len])
        plt.xlabel('Flow volume (Normalized)')
        plt.ylabel('Power')
        plt.title('FVC = {:.2f} L'.format(fvc))

        # Set limits
        if len(flow_volume) > 0:
            x_limit = max(np.abs(flow_volume)) * np.array([-1.1, 1.1])
            plt.xlim(x_limit)
            
        if len(flow_rate_smoothed) > 0:
            y_limit = max(np.abs(flow_rate_smoothed)) * np.array([-1.1, 1.1])
            plt.ylim(y_limit)

        # Add zero lines
        plt.axhline(0, color='k', linestyle='--')
        plt.axvline(0, color='k', linestyle='--')

        # Save the second plot
        img_bytes_io_2 = io.BytesIO()
        plt.savefig(img_bytes_io_2)
        img_bytes_io_2.seek(0)
        img_base64_2 = base64.b64encode(img_bytes_io_2.read()).decode('utf-8')
        plt.close()
        
    except Exception as e:
        # If there's an error in the flow volume calculation, create an error plot
        plt.figure()
        plt.text(0.5, 0.5, f"Error in FVC calculation: {str(e)}", 
                horizontalalignment='center', verticalalignment='center',
                transform=plt.gca().transAxes, fontsize=14, color='red')
        plt.xlabel('Flow Volume')
        plt.ylabel('Power')
        plt.grid(True)
        
        # Save the error plot
        img_bytes_io_2 = io.BytesIO()
        plt.savefig(img_bytes_io_2)
        img_bytes_io_2.seek(0)
        img_base64_2 = base64.b64encode(img_bytes_io_2.read()).decode('utf-8')
        plt.close()

    return img_base64_1, img_base64_2

def svc_calc(data):
    imvmin,mvmin,imvmax,mvmax,mvAvgPow,mvAvgX,mvAvgY,master_data = utli_fun(data)

    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmax / 1000, mvmax, 'or')
    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmin / 1000, mvmin, 'ob')
    plt.xlabel('Time (s)')
    plt.ylabel('Pow')
    plt.grid(True)
    plt.minorticks_on()
    img_bytes_io = io.BytesIO()
    plt.savefig(img_bytes_io)
    #f_name = fol+'/'+bacteria+'_'+d_type+".png"
    #plt.savefig(f_name)
    img_bytes_io.seek(0)
    img_base64 = base64.b64encode(img_bytes_io.read()).decode('utf-8')
    plt.close()  # Close the plot to free up resources
    return img_base64

def mvv_calc(data):
    imvmin,mvmin,imvmax,mvmax,mvAvgPow,mvAvgX,mvAvgY,master_data = utli_fun(data)

    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmax / 1000, mvmax, 'or')
    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmin / 1000, mvmin, 'ob')
    plt.xlabel('Time (s)')
    plt.ylabel('Pow')
    plt.grid(True)
    plt.minorticks_on()
    img_bytes_io = io.BytesIO()
    plt.savefig(img_bytes_io)
    #f_name = fol+'/'+bacteria+'_'+d_type+".png"
    #plt.savefig(f_name)
    img_bytes_io.seek(0)
    img_base64 = base64.b64encode(img_bytes_io.read()).decode('utf-8')
    plt.close()  # Close the plot to free up resources
    return img_base64

def sm_calc(df):
    
    imvmin,mvmin,imvmax,mvmax,mvAvgPow,mvAvgX,mvAvgY,master_data = utli_fun(df)

    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmax / 1000, mvmax, 'or')
    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmin / 1000, mvmin, 'ob')
    plt.xlabel('Time (s)')
    plt.ylabel('Pow')
    plt.grid(True)
    plt.minorticks_on()

    i = 0
    while (i <len(imvmax)-1) & (i <len(imvmin)-1) :
        
        #if i<=(len(imvmax)-1) & i<=(len(imvmin)-1):
        if imvmax[0] > imvmin[0]:
            #print("i is: ",i," and Imvmax : ",len(imvmin)-1)
            if imvmax[i] > imvmin[i+1]:
                imvmin = np.delete(imvmin, i)
                i = 0
                continue
            if imvmax[i] < imvmin[i]:
                imvmax = np.delete(imvmax, i)
                i = 0
                continue
            if i == len(imvmax):
                break
            i += 1
        if imvmin[0] > imvmax[0]:
            #print("i is: ",i," and Imvmax : ",len(imvmax)-1)
            if imvmax[i] > imvmin[i]:
                imvmin = np.delete(imvmin, i)
                i = 0
                continue
            if i < len(imvmax):
                if imvmax[i+1] < imvmin[i]:
                    imvmax = np.delete(imvmax, i)
                    i = 0
                    continue
            if i == len(imvmin):
                break
            i += 1
    Inhalationn = np.zeros(len(imvmin) - 1)
    InhalationVol = np.zeros(len(imvmin) - 1)
    Exhalationn = np.zeros(len(imvmin) - 1)
    ExhalationVol = np.zeros(len(imvmin) - 1)
    volarr = np.zeros(len(imvmin) - 1)
    
    if len(imvmax) >= len(imvmin):
        print(imvmin[0] > imvmax[0])
        if imvmin[0] > imvmax[0]:
            for i in range(len(imvmin) - 1):
                Inhalationn[i] = imvmax[i + 1] - imvmin[i]
                InhalationVol[i] = np.trapz(mvAvgPow[imvmin[i]:imvmax[i+1]])
                Exhalationn[i] = imvmin[i] - imvmax[i]
                ExhalationVol[i] = np.trapz(mvAvgPow[imvmax[i]:imvmin[i]])
    else:
        for i in range(len(imvmax) - 1):
            Inhalationn[i] = imvmax[i] - imvmin[i + 1]
            InhalationVol[i] = np.trapz(mvAvgPow[imvmin[i + 1]:imvmax[i]])
            Exhalationn[i] = imvmin[i + 1] - imvmax[i]
            ExhalationVol[i] = np.trapz(mvAvgPow[imvmax[i]:imvmin[i + 1]])

    Inhalation_Duration = sum(Inhalationn)
    Exhalation_Duration = sum(Exhalationn)
    Inhalation_Volume = sum(InhalationVol)
    Exhalation_Volume = sum(ExhalationVol)
    Resp_Rate = round(0.5 * (len(imvmax) + len(imvmin)) * 60 / 20)
    Vel_x = np.diff(mvAvgX)
    Vel_y = np.diff(mvAvgY)
    MagnitudeParameters_Vel = np.sqrt(Vel_x**2 + Vel_y**2)

    accel_x = np.diff(np.diff(mvAvgX))
    accel_y = np.diff(np.diff(mvAvgY))
    MagnitudeParameters_accel = np.sqrt(accel_x**2 + accel_y**2)

    dif_of_accel_x = np.diff(np.diff(np.diff(mvAvgX)))
    dif_of_accel_y = np.diff(np.diff(np.diff(mvAvgY)))
    MagnitudeParameters_dif_of_accel = np.sqrt(dif_of_accel_x**2 + dif_of_accel_y**2)

    exvol0_real = ExhalationVol[0]
    exvol0_Normalized = ExhalationVol[0] / sum(ExhalationVol)
    LungCapacity = ((Exhalation_Volume) / len(ExhalationVol)) * 5 * 10**-4

    if LungCapacity >= 6:
        LungCapacity_Qual = 1
    else:
        LungCapacity_Qual = LungCapacity / 6

    if Resp_Rate >= 12 and Resp_Rate <= 20:
        Quality = 1
    else:
        if Resp_Rate > 20:
            Quality = 1 - (Resp_Rate - 20) / 20
        else:
            Quality = 1 - (12 - Resp_Rate) / 12
    QualityThreshold1 = 0.7
    QualityThreshold2 = 0.5
    if LungCapacity_Qual < 0.6:
        if Quality < QualityThreshold1 and Quality > QualityThreshold2:
            FinalQI = 2
        elif Quality <= QualityThreshold2:
            FinalQI = 3
        else:
            FinalQI = 1
    else:
        FinalQI = 1

    return Resp_Rate, LungCapacity, Quality

def raw_data_calc(data):
    
    imvmin,mvmin,imvmax,mvmax,mvAvgPow,mvAvgX,mvAvgY,master_data = utli_fun(data)

    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmax / 1000, mvmax, 'or')
    plt.plot(np.arange(len(mvAvgPow)) / 1000, mvAvgPow, '-k', imvmin / 1000, mvmin, 'ob')
    plt.xlabel('Time (s)')
    plt.ylabel('Pow')
    plt.grid(True)
    plt.minorticks_on()
    img_bytes_io = io.BytesIO()
    plt.savefig(img_bytes_io)
    #f_name = fol+'/'+bacteria+'_'+d_type+".png"
    #plt.savefig(f_name)
    img_bytes_io.seek(0)
    img_base64 = base64.b64encode(img_bytes_io.read()).decode('utf-8')
    plt.close()  # Close the plot to free up resources
    return img_base64
    #return "raw_data"

def utli_fun(df):
    mv_avg_window=1000
    master_data = df  # Assuming CSV format with header row
    master_data.dropna(inplace=True)

    mvAvgX = master_data['Xpos'].rolling(window=mv_avg_window,min_periods=1).mean()
    mvAvgY = master_data['Ypos'].rolling(window=mv_avg_window,min_periods=1).mean()
    mvAvgPow = master_data['Pow'].rolling(window=mv_avg_window,min_periods=1).mean()
    x=mvAvgPow
    # Find peaks
    imvmax, _ = find_peaks(x,prominence=0.02)

    # Find valleys (negative peaks)
    imvmin, _ = find_peaks(-x,prominence=0.05)
    mvmax = mvAvgPow[imvmax]
    mvmin = mvAvgPow[imvmin]
    imvmax1 = imvmax
    imvmin1 = imvmin
    mvmax1 = mvmax
    mvmin1 = mvmin

    return imvmin,mvmin,imvmax,mvmax,mvAvgPow,mvAvgX,mvAvgY,master_data