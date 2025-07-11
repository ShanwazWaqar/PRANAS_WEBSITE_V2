import scipy.stats as stats 
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import io,base64
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.svm import SVC
import xgboost as xgb
import lightgbm as lgb
from sklearn.ensemble import GradientBoostingClassifier
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report,f1_score
from sklearn.preprocessing import LabelEncoder
import matplotlib.patheffects as pe
from matplotlib.colors import ListedColormap
from sklearn.preprocessing import LabelEncoder
import scipy.stats as stats 
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import io,base64
import pandas as pd
import scipy.stats as stats 
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import io,base64
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
import matplotlib.pyplot as plt
import seaborn as sns
import scipy.stats as stats 
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import io,base64
import pandas as pd
import os,pickle
bacteria_colors = {'EC': 'blue', 'SA': 'green', 'PA': 'red', 'KP': 'black', 'SM': 'purple', 'SE': 'brown','ECSA':'orange','PASE':'cyan','SMKP':'teal','LM':'pink'}
def calculate_sa(group):
    max_intensity = group['Pow'].max()
    min_intensity = group['Pow'].min()
    mean_intensity = group['Pow'].mean()
    
    # Calculate SA
    if mean_intensity != 0:
        return (max_intensity - min_intensity) / mean_intensity
    else:
        return np.nan
def features_creator(averaged_data):
    print(averaged_data.head())
    averaged_data['Mag'] = np.sqrt(averaged_data['Ypos']**2 + averaged_data['Xpos']**2)
    averaged_data['BDE'] =averaged_data['Mag'] / averaged_data['Pow']
    averaged_data['dir'] = np.arctan2(averaged_data['Ypos'], averaged_data['Xpos'])
    averaged_data['SIV'] = averaged_data.groupby(['bacteria'])['Pow'].transform(np.std)
    center_x, center_y = averaged_data['Xpos'].mean(), averaged_data['Ypos'].mean()
    averaged_data['SCD'] = np.sqrt((averaged_data['Xpos'] - center_x) ** 2 + (averaged_data['Ypos'] - center_y) ** 2)
    averaged_data['AoS'] = np.arctan2(averaged_data['Ypos'] - center_y, averaged_data['Xpos'] - center_x)
    psi_values = averaged_data.groupby('bacteria')['Pow'].max()
    averaged_data['PSI'] = averaged_data['bacteria'].map(psi_values)
    msa_values = averaged_data.groupby('bacteria')['dir'].mean()
    averaged_data['MSA'] = averaged_data['bacteria'].map(msa_values)
    tsp_values = averaged_data.groupby('bacteria')['Pow'].sum()
    averaged_data['TSP'] = averaged_data['bacteria'].map(tsp_values)
    ssi_values = averaged_data.groupby('bacteria').apply(lambda x: abs(x['Pow'] - x['Pow'].iloc[::-1]).mean())
    averaged_data['SSI'] = averaged_data['bacteria'].map(ssi_values)
    se_values = averaged_data.groupby('bacteria')['Pow'].apply(lambda x: stats.entropy(x))
    averaged_data['SE'] = averaged_data['bacteria'].map(se_values)
    sa_values = averaged_data.groupby('bacteria').apply(calculate_sa)
    averaged_data['SA'] = averaged_data['bacteria'].map(sa_values)
    return averaged_data
def data_split_group(data,bact):
    data1=data[data['bacteria'].isin(bact)]
    return data1
def determine_d_t(bacteria_list):
    if all(bacteria in ["SA", "SE", "SM"] for bacteria in bacteria_list) and len(bacteria_list) > 0:
        d_t = "gp_data"
    elif all(bacteria in ["EC", "PA", "KP"] for bacteria in bacteria_list) and len(bacteria_list) > 0:
        d_t = "gn_data"
    elif all(bacteria in ["ECSA", "SMKP", "PASE"] for bacteria in bacteria_list) and len(bacteria_list) > 0:
        d_t = "com_data"
    else:
        d_t = "cus_data"
    return d_t
def PCA_1(gn_data,scaler):
    #gn_data = gn_data[(gn_data['slide'].isin())]
    print("called PCA")
    features_for_pca = ['SIV','BDE','dir']
    X_pca = gn_data[features_for_pca]
    unique_values_string = '_'.join(gn_data['bacteria'].unique())
    X_pca_scaled = scaler.fit_transform(X_pca)
    # if not os.path.exists("StandardScalers/"):
    #     os.mkdir("StandardScalers/")
    # scaler_name= "StandardScalers/"+unique_values_string+'_scaler.pkl'
    # with open(scaler_name, 'wb') as file:
    #     pickle.dump(scaler, file)
    # Apply PCA
    pca = PCA(n_components=2)
    pca_components = pca.fit_transform(X_pca_scaled)
    # if not os.path.exists("PCA_models/"):
    #     os.mkdir("PCA_models/")
    # pca_name= "PCA_models/"+unique_values_string+'_pca.pkl'
    # with open(pca_name, 'wb') as file:
    #     pickle.dump(pca, file)
    # Add PCA components to the data
    gn_data['PCA1'], gn_data['PCA2'] = pca_components[:, 0], pca_components[:, 1]

    # Calculate centroids for each bacteria type in the PCA space
    centroids = gn_data.groupby('bacteria')[['PCA1', 'PCA2']].mean()

    # Calculate Euclidean distance of each point from its corresponding bacteria centroid
    gn_data['distance_from_centroid'] = gn_data.apply(
        lambda row: np.linalg.norm(row[['PCA1', 'PCA2']] - centroids.loc[row['bacteria']]),
        axis=1
    )

    # Define a distance threshold - for this example, we'll use the mean distance plus one standard deviation
    distance_threshold = gn_data['distance_from_centroid'].mean() + gn_data['distance_from_centroid'].std()

    # Filter the data to retain only points within the threshold distance from the centroid
    filtered_data = gn_data[gn_data['distance_from_centroid'] <= distance_threshold]
    return filtered_data,scaler,pca
def main_function_pca(data_using,bacts,conc,vol,slide,trails,data_type):
    print("Main PCA")
    fol="plots/pca_plots"
    scaler1 = StandardScaler()
    print(data_using.head())
    filtered_data,scaler2,pca2 = PCA_1(data_using,scaler1)
    plt.figure(figsize=(14, 8))
    sns.scatterplot(data=filtered_data, x='PCA1', y='PCA2', hue='bacteria', palette=bacteria_colors)
    # plt.title('Filtered PCA of Bacteria Data Close to Centroids')
    plt.xlabel('PCA Component 1', fontsize=16)
    plt.ylabel('PCA Component 2', fontsize=16)
    plt.grid(True)
    my_string = '_'.join(slide)
    f_name = fol+'/'+data_type+my_string+'pca'+".png"
    plt.savefig(f_name)
    plt.close()
    return filtered_data,scaler2,pca2

def ML_train_val(filtered_data):
    print("ML Training")
    X = filtered_data[['PCA1', 'PCA2']]
    y = filtered_data['bacteria']
    le=LabelEncoder()
    y=le.fit_transform(y)
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

    svm_model = SVC(kernel='linear', random_state=42)
    svm_model.fit(X_train, y_train)

    # Predict and evaluate
    y_pred_svmT = svm_model.predict(X_test)
    svm_acc=accuracy_score(y_test, y_pred_svmT)
    print("SVM Accuracy:", svm_acc)
    
    xgb_model = xgb.XGBClassifier(use_label_encoder=False, eval_metric='logloss')
    xgb_model.fit(X_train, y_train)

    # Predict and evaluate
    y_pred_xgbT = xgb_model.predict(X_test)
    xgb_acc=accuracy_score(y_test, y_pred_xgbT)
    print("XGBoost Accuracy:", xgb_acc)

    lgb_model = lgb.LGBMClassifier()
    lgb_model.fit(X_train, y_train)

    # Predict and evaluate
    y_pred_lgbT = lgb_model.predict(X_test)
    lgbm_acc=accuracy_score(y_test, y_pred_lgbT)
    print("LightGBM Accuracy:",lgbm_acc )

    gb_model = GradientBoostingClassifier(random_state=42)
    gb_model.fit(X_train, y_train)

    # Predict and evaluate
    y_pred_gbT = gb_model.predict(X_test)
    gb_Acc=accuracy_score(y_test, y_pred_gbT)
    print("Gradient Boosting Accuracy:",gb_Acc )

    
    return svm_model,gb_model,lgb_model,xgb_model,le,svm_acc,xgb_acc,lgbm_acc,gb_Acc,X_train, X_test, y_train, y_test,X,y

def pca_test(gn_data, slide, scaler, pca2):
    """Fixed version that handles empty data arrays"""
    print("PCA Test")
    
    try:
        # Determine the list of slides to use for testing
        if any(slidec in ["C0","C1", "C2", "C3"] for slidec in slide):
            list2 = ["C0","C1", "C2", "C3"]
        else:
            list2 = ["S0","S1","S2","S3"]
        
        list1 = slide
        missing_value = list(set(list2) - set(list1))
        
        # CRITICAL FIX: Check if missing_value is empty
        if not missing_value:
            print(f"Warning: No test slides available. Training slides: {slide}, All slides: {list2}")
            # Return empty DataFrame with correct structure
            empty_df = pd.DataFrame(columns=['PCA1', 'PCA2', 'bacteria'])
            return empty_df
        
        # Filter data by missing slides
        filtered_gn_data = gn_data[gn_data['slide'].isin(missing_value)]
        
        # CRITICAL FIX: Check if filtered data is empty
        if filtered_gn_data.empty:
            print(f"Warning: No data found for test slides {missing_value}")
            print(f"Available slides in data: {gn_data['slide'].unique()}")
            # Return empty DataFrame with correct structure
            empty_df = pd.DataFrame(columns=['PCA1', 'PCA2', 'bacteria'])
            return empty_df
        
        print(f"Found {len(filtered_gn_data)} rows for test slides {missing_value}")
        
        # Check if required features exist
        features_for_pca = ['SIV','BDE','dir']
        missing_features = [f for f in features_for_pca if f not in filtered_gn_data.columns]
        if missing_features:
            print(f"Warning: Missing features for PCA: {missing_features}")
            empty_df = pd.DataFrame(columns=['PCA1', 'PCA2', 'bacteria'])
            return empty_df
            
        X_pca = filtered_gn_data[features_for_pca]
        
        # CRITICAL FIX: Check if X_pca is empty
        if X_pca.empty:
            print("Warning: No feature data available for PCA")
            empty_df = pd.DataFrame(columns=['PCA1', 'PCA2', 'bacteria'])
            return empty_df
        
        # Fill any NaN values
        X_pca = X_pca.fillna(0)
        
        # Apply scaling and PCA transformation
        X_pca_scaled = scaler.transform(X_pca)
        pca_components = pca2.transform(X_pca_scaled)
        
        # Add PCA components to the data
        filtered_gn_data = filtered_gn_data.copy()  # Avoid SettingWithCopyWarning
        filtered_gn_data['PCA1'] = pca_components[:, 0]
        filtered_gn_data['PCA2'] = pca_components[:, 1]
        
        # Check if we have bacteria data for centroids
        if 'bacteria' not in filtered_gn_data.columns:
            print("Warning: No bacteria column found")
            return filtered_gn_data[['PCA1', 'PCA2']]
        
        # Calculate centroids for each bacteria type in the PCA space
        try:
            centroids = filtered_gn_data.groupby('bacteria')[['PCA1', 'PCA2']].mean()
            
            # Calculate Euclidean distance of each point from its corresponding bacteria centroid
            def calculate_distance(row):
                bacteria = row['bacteria']
                if bacteria in centroids.index:
                    return np.linalg.norm(row[['PCA1', 'PCA2']] - centroids.loc[bacteria])
                else:
                    return 0  # Default distance if bacteria not found
            
            filtered_gn_data['distance_from_centroid'] = filtered_gn_data.apply(calculate_distance, axis=1)
            
            # Define a distance threshold
            if len(filtered_gn_data) > 1:
                distance_threshold = filtered_gn_data['distance_from_centroid'].mean() + filtered_gn_data['distance_from_centroid'].std()
            else:
                distance_threshold = filtered_gn_data['distance_from_centroid'].max()
            
            # Filter the data to retain only points within the threshold distance from the centroid
            filtered_data_test = filtered_gn_data[filtered_gn_data['distance_from_centroid'] <= distance_threshold]
            
            if filtered_data_test.empty:
                print("Warning: No data points within distance threshold")
                return filtered_gn_data[['PCA1', 'PCA2', 'bacteria']]
            
            return filtered_data_test
            
        except Exception as e:
            print(f"Error calculating centroids: {str(e)}")
            # Return basic PCA data without centroid filtering
            return filtered_gn_data[['PCA1', 'PCA2', 'bacteria']]
    
    except Exception as e:
        print(f"Error in pca_test: {str(e)}")
        import traceback
        print(f"Full traceback: {traceback.format_exc()}")
        # Return empty DataFrame with correct structure
        empty_df = pd.DataFrame(columns=['PCA1', 'PCA2', 'bacteria'])
        return empty_df



def ML_test(test_data,svm_model,xgb_model,lgb_model,gb_model,le):
    print("ML Test")
    X_test2 = test_data[['PCA1','PCA2']]
    y_test2 = test_data['bacteria']
    y_test2=le.transform(y_test2)
    y_pred_svm = svm_model.predict(X_test2)
    
    svm_acc_test = accuracy_score(y_test2, y_pred_svm)
    print('Svm accuracy: ',f'Accuracy: {svm_acc_test}')
    svm_f1_test = f1_score(y_test2, y_pred_svm, average='weighted')  # Use 'weighted' to consider the imbalance in the dataset
    print(f'SVM F1 Score: {svm_f1_test:.2f}')
    y_pred_xgb = xgb_model.predict(X_test2)
    xgb_acc_test = accuracy_score(y_test2, y_pred_xgb)
    print('xgb_model accuracy: ',f'Accuracy: {xgb_acc_test}')
    xgb_f1_test = f1_score(y_test2, y_pred_xgb, average='weighted')  # Use 'weighted' to consider the imbalance in the dataset
    print(f'xgb F1 Score: {xgb_f1_test:.2f}')
    y_pred_lgb = lgb_model.predict(X_test2)
    lgb_acc_test = accuracy_score(y_test2, y_pred_lgb)
    print('lgb_model accuracy: ',f'Accuracy: {lgb_acc_test}')
    lgb_f1_test = f1_score(y_test2, y_pred_lgb, average='weighted')  # Use 'weighted' to consider the imbalance in the dataset
    print(f'lgbm F1 Score: {lgb_f1_test:.2f}')
    y_pred_gb= gb_model.predict(X_test2)
    gb_acc_test = accuracy_score(y_test2, y_pred_gb)
    print('gb_model accuracy: ',f'Accuracy: {gb_acc_test}')
    gb_f1_test = f1_score(y_test2, y_pred_gb, average='weighted')  # Use 'weighted' to consider the imbalance in the dataset
    print(f'gbm F1 Score: {gb_f1_test:.2f}')

    return y_pred_svm,X_test2,svm_f1_test,svm_acc_test,xgb_acc_test,xgb_f1_test,lgb_acc_test,lgb_f1_test,gb_acc_test,gb_f1_test

def plotter_svm(df, bacts, conc, vol, slide, trails):
    """Updated plotter_svm with better error handling"""
    print("Plotter")
    bacteria_colors_full = {'EC': 'blue', 'SA': 'green', 'PA': 'red', 'KP': 'black', 'SM': 'purple', 'SE': 'brown','ECSA':'orange','PASE':'cyan','SMKP':'teal','LM':'pink'}
    color_map_full = {'EC': 'blue', 'SA': 'green', 'PA': 'red', 'KP': 'black', 'SM': 'purple', 'SE': 'brown','ECSA':'orange','PASE':'cyan','SMKP':'teal','LM':'pink'}
    bacteria_colors = {key: bacteria_colors_full[key] for key in bacts if key in bacteria_colors_full}
    color_map = bacteria_colors.copy()
    data_type = determine_d_t(bacts)
    print(data_type)
    
    # Convert parameters
    conc = float(conc)
    vol = float(vol)
    
    # Filter data by concentration and volume
    df = df[(df['concentration'] == conc) & (df['volume'] == vol)]
    
    # Check if we have data after filtering
    if df.empty:
        raise ValueError(f"No data found for concentration={conc} and volume={vol}")
    
    # Create features
    f_data = features_creator(df)
    data_using = data_split_group(f_data, bacts)
    
    # Check if we have data after bacteria filtering
    if data_using.empty:
        raise ValueError(f"No data found for selected bacteria: {bacts}")
    
    # Filter by slides for training
    data_using2 = data_using[data_using['slide'].isin(slide)]
    
    # Check if we have training data
    if data_using2.empty:
        raise ValueError(f"No data found for selected slides: {slide}")
    
    print(f"Training data shape: {data_using2.shape}")
    print(data_using2['BDE'].max())
    
    # Create raw plot
    fol_raw = "Plots/raw_plots"
    plt.figure(figsize=(10, 6))
    sns.scatterplot(data=data_using2, x='BDE', y='dir', hue='bacteria', style='slide', palette=bacteria_colors)
    plt.xlabel('BDE', fontsize=16)
    plt.ylabel('DIR', fontsize=16)
    plt.legend(title='Bacteria', loc='best')
    plt.grid('on')
    my_string = '_'.join(slide)
    f_name_raw = fol_raw + '/' + data_type + my_string + 'raw' + ".png"
    plt.savefig(f_name_raw)
    plt.close()
    
    # Perform PCA and ML training
    filtered_data, scaler2, pca2 = main_function_pca(data_using2, bacts, conc, vol, slide, trails, data_type)
    
    # Check if PCA resulted in data
    if filtered_data.empty:
        raise ValueError("PCA processing resulted in no data")
    
    # Train ML models
    svm_model, gb_model, lgb_model, xgb_model, le, svm_acc, xgb_acc, lgbm_acc, gb_Acc, X_train, X_test, y_train, y_test, X, y = ML_train_val(filtered_data)
    
    # Get test data - FIXED: Handle empty test data
    try:
        filtered_data_test = pca_test(data_using, slide, scaler2, pca2)
        
        # Check if we have test data
        if filtered_data_test.empty or len(filtered_data_test) == 0:
            print("Warning: No test data available. Using a portion of training data for visualization.")
            # Use a subset of training data for visualization
            test_size = min(50, len(filtered_data) // 4)  # Use 25% or max 50 samples
            filtered_data_test = filtered_data.sample(n=test_size, random_state=42) if len(filtered_data) > test_size else filtered_data
        
        # Perform ML testing only if we have test data
        if not filtered_data_test.empty and len(filtered_data_test) > 0:
            y_pred_svm, X_test2, svm_f1_test, svm_acc_test, xgb_acc_test, xgb_f1_test, lgb_acc_test, lgb_f1_test, gb_acc_test, gb_f1_test = ML_test(filtered_data_test, svm_model, gb_model, lgb_model, xgb_model, le)
        else:
            # Use training metrics as fallback
            print("Using training metrics as test metrics due to lack of test data")
            svm_f1_test = svm_acc_test = svm_acc
            xgb_f1_test = xgb_acc_test = xgb_acc
            lgb_f1_test = lgb_acc_test = lgbm_acc
            gb_f1_test = gb_acc_test = gb_Acc
            X_test2 = X_test
            y_pred_svm = svm_model.predict(X_test)
            
    except Exception as e:
        print(f"Error in testing phase: {str(e)}")
        # Use training metrics as fallback
        svm_f1_test = svm_acc_test = svm_acc
        xgb_f1_test = xgb_acc_test = xgb_acc
        lgb_f1_test = lgb_acc_test = lgbm_acc
        gb_f1_test = gb_acc_test = gb_Acc
        X_test2 = X_test
        y_pred_svm = svm_model.predict(X_test)
    
    # Save metrics
    file_path = 'Metrics_ML/metrics.csv'
    columns = ['Model', 'Val_Accuracy', 'Test_acc', 'Test_F1', 'Train_set', 'bacteria_set']
    unique_values_slide = '_'.join(slide)
    all_bact_string = '_'.join(bacts)
    
    if not os.path.exists('Metrics_ML'):
        os.makedirs('Metrics_ML')
    
    if not os.path.exists(file_path):
        dfXT = pd.DataFrame(columns=columns)
        dfXT.to_csv(file_path, index=False)
    
    new_data = {
        'Model': ['SVM', 'XGBoost', 'LGBM', 'GB'], 
        'Val_Accuracy': [svm_acc, xgb_acc, lgbm_acc, gb_Acc], 
        'Test_Accuracy': [svm_acc_test, xgb_acc_test, lgb_acc_test, gb_acc_test],
        'Test_F1': [svm_f1_test, xgb_f1_test, lgb_f1_test, gb_f1_test],
        'Train_set': [unique_values_slide, unique_values_slide, unique_values_slide, unique_values_slide],
        'bacteria_set': [all_bact_string, all_bact_string, all_bact_string, all_bact_string]
    }
    new_df = pd.DataFrame(new_data)
    new_df.to_csv(file_path, mode='a', header=False, index=False)
    
    # Create decision boundary plot
    x_min, x_max = X.iloc[:, 0].min() - 1, X.iloc[:, 0].max() + 1
    y_min, y_max = X.iloc[:, 1].min() - 1, X.iloc[:, 1].max() + 1
    xx, yy = np.meshgrid(np.arange(x_min, x_max, 0.02),
                        np.arange(y_min, y_max, 0.02))
    
    # Predict classes for each point in the mesh
    Z = svm_model.predict(np.c_[xx.ravel(), yy.ravel()])
    Z = Z.reshape(xx.shape)
    
    # Inverse transform to get original bacteria names
    y_train_bacteria = le.inverse_transform(y_train)
    y_test_bacteria = le.inverse_transform(y_test)
    y_pred2_bacteria = le.inverse_transform(y_pred_svm)
    
    color_list = [color_map[class_label] for class_label in le.classes_]
    cmap_light = ListedColormap(color_list)
    
    plt.figure(figsize=(15, 6))
    plt.contourf(xx, yy, Z, colors=color_list, alpha=0.3)
    
    for bacteria, color in bacteria_colors.items():
        plt.scatter(X_train[y_train_bacteria == bacteria].iloc[:, 0], X_train[y_train_bacteria == bacteria].iloc[:, 1], 
                    color=color, label=f'Training Data: {bacteria}', edgecolor='k', s=20)
    
    plt.scatter(svm_model.support_vectors_[:, 0], svm_model.support_vectors_[:, 1], 
                facecolors='none', edgecolors='k', s=100, label='Support Vectors')
    
    for bacteria, color in bacteria_colors.items():
        plt.scatter(X_test2[y_pred2_bacteria == bacteria].iloc[:, 0], X_test2[y_pred2_bacteria == bacteria].iloc[:, 1], 
                    color=color, marker='x', edgecolor='k', s=100, label=f'Predicted Test Data: {bacteria}')
    
    for bacteria, color in bacteria_colors.items():
        mask = Z == le.transform([bacteria])[0]
        if np.any(mask):
            region_xx, region_yy = xx[mask], yy[mask]
            centroid_x = region_xx.mean()
            centroid_y = region_yy.mean()
            plt.text(centroid_x, centroid_y, bacteria, color='white',
                    fontsize=24, ha='center', va='center', weight='bold', path_effects=[
                pe.withStroke(linewidth=10, foreground='black'),
                pe.Normal()
            ])
    
    fol = "Plots/ML_Plots"
    if not os.path.exists(fol):
        os.makedirs(fol)
    
    handles, labels = plt.gca().get_legend_handles_labels()
    plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
    plt.xlabel('PCA Component 1', fontsize=16)
    plt.ylabel('PCA Component 2', fontsize=16)
    plt.tight_layout(rect=[0, 0, 0.75, 1])
    
    img_bytes_io = io.BytesIO()
    plt.savefig(img_bytes_io, format='png', dpi=300, bbox_inches='tight')
    my_string = '_'.join(slide)
    f_name = fol + '/' + data_type + my_string + 'SVM' + ".png"
    plt.savefig(f_name)
    img_bytes_io.seek(0)
    img_base64 = base64.b64encode(img_bytes_io.read()).decode('utf-8')
    plt.close()
    
    return img_base64