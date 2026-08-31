[1mdiff --git a/.gitignore b/.gitignore[m
[1mindex 737faff..00a6912 100644[m
[1m--- a/.gitignore[m
[1m+++ b/.gitignore[m
[36m@@ -1,33 +1,33 @@[m
[31m-# Archivos de autenticacion - NO SUBIR[m
[31m-credentials.json[m
[31m-token.pickle[m
[31m-cuit_mapping.json[m
[31m-[m
[31m-# Archivos temporales[m
[31m-temp/[m
[31m-*.tmp[m
[31m-[m
[31m-# Python[m
[31m-__pycache__/[m
[31m-*.py[cod][m
[31m-*$py.class[m
[31m-*.so[m
[31m-.Python[m
[31m-env/[m
[31m-venv/[m
[31m-*.egg-info/[m
[31m-dist/[m
[31m-build/[m
[31m-[m
[31m-# IDEs[m
[31m-.vscode/[m
[31m-.idea/[m
[31m-*.swp[m
[31m-*.swo[m
[31m-[m
[31m-# OS[m
[31m-.DS_Store[m
[31m-Thumbs.db[m
[31m-[m
[31m-# Logs[m
[31m-*.log[m
[32m+[m[32m# Archivos de autenticacion - NO SUBIR[m[41m[m
[32m+[m[32mcredentials.json[m[41m[m
[32m+[m[32mtoken.pickle[m[41m[m
[32m+[m[32mcuit_mapping.json[m[41m[m
[32m+[m[41m[m
[32m+[m[32m# Archivos temporales[m[41m[m
[32m+[m[32mtemp/[m[41m[m
[32m+[m[32m*.tmp[m[41m[m
[32m+[m[41m[m
[32m+[m[32m# Python[m[41m[m
[32m+[m[32m__pycache__/[m[41m[m
[32m+[m[32m*.py[cod][m[41m[m
[32m+[m[32m*$py.class[m[41m[m
[32m+[m[32m*.so[m[41m[m
[32m+[m[32m.Python[m[41m[m
[32m+[m[32menv/[m[41m[m
[32m+[m[32mvenv/[m[41m[m
[32m+[m[32m*.egg-info/[m[41m[m
[32m+[m[32mdist/[m[41m[m
[32m+[m[32mbuild/[m[41m[m
[32m+[m[41m[m
[32m+[m[32m# IDEs[m[41m[m
[32m+[m[32m.vscode/[m[41m[m
[32m+[m[32m.idea/[m[41m[m
[32m+[m[32m*.swp[m[41m[m
[32m+[m[32m*.swo[m[41m[m
[32m+[m[41m[m
[32m+[m[32m# OS[m[41m[m
[32m+[m[32m.DS_Store[m[41m[m
[32m+[m[32mThumbs.db[m[41m[m
[32m+[m[41m[m
[32m+[m[32m# Logs[m[41m[m
[32m+[m[32m*.log[m[41m[m
[1mdiff --git a/Utils/__init__.py b/Utils/__init__.py[m
[1mindex dd7ee44..0720fdf 100644[m
[1m--- a/Utils/__init__.py[m
[1m+++ b/Utils/__init__.py[m
[36m@@ -1 +1 @@[m
[31m-# Utils package[m
[32m+[m[32m# Utils package[m[41m[m
[1mdiff --git a/Utils/excel_processor.py b/Utils/excel_processor.py[m
[1mindex f04bc98..444b7be 100644[m
[1m--- a/Utils/excel_processor.py[m
[1m+++ b/Utils/excel_processor.py[m
[36m@@ -143,6 +143,10 @@[m [mclass ExcelProcessor:[m
             raise ValueError("No se encontraron fechas válidas en el archivo")[m
 [m
         # Encontrar el mes más frecuente[m
[32m+[m[32m        # Le indicamos que el DÍA va primero (formato DD/MM/AAAA de AFIP)[m
[32m+[m[32m        df_temp['Fecha'] = pd.to_datetime(df_temp['Fecha'], dayfirst=True, errors='coerce')[m
[32m+[m
[32m+[m[32m        # Ahora el conteo de meses va a ser el correcto[m
         month_counts = df_temp['Fecha'].dt.month.value_counts()[m
         most_common_month = month_counts.idxmax()[m
 [m
[36m@@ -223,7 +227,10 @@[m [mclass ExcelProcessor:[m
             # Ya es string, no hacer nada[m
             pass[m
         else:[m
[31m-            # Es datetime, convertir a string[m
[32m+[m[32m            # 1. Aseguramos que Pandas entienda que son fechas con formato de nuestro país[m
[32m+[m[32m            df_clean['Fecha'] = pd.to_datetime(df_clean['Fecha'], dayfirst=True, errors='coerce')[m
[32m+[m
[32m+[m[32m            # 2. Ahora sí aplicamos el formato visual de manera segura (esta es tu línea original)[m
             df_clean['Fecha'] = df_clean['Fecha'].dt.strftime('%d/%m/%Y')[m
         [m
         # Calcular totales para las columnas numéricas[m
[1mdiff --git a/requirements.txt b/requirements.txt[m
[1mindex d84b420..89a6273 100644[m
[1m--- a/requirements.txt[m
[1m+++ b/requirements.txt[m
[36m@@ -1,10 +1,10 @@[m
[31m-fastapi==0.104.1[m
[31m-uvicorn[standard]==0.24.0[m
[31m-python-multipart==0.0.6[m
[31m-pandas==2.1.3[m
[31m-openpyxl==3.1.2[m
[31m-google-auth==2.23.4[m
[31m-google-auth-oauthlib==1.1.0[m
[31m-google-auth-httplib2==0.1.1[m
[31m-google-api-python-client==2.108.0[m
[31m-tabulate==0.9.0[m
[32m+[m[32mfastapi #==0.104.1[m
[32m+[m[32muvicorn #[standard]#==0.24.0[m
[32m+[m[32mpython-multipart #==0.0.6[m
[32m+[m[32mpandas #==2.1.3[m
[32m+[m[32mopenpyxl #==3.1.2[m
[32m+[m[32mgoogle-auth #==2.23.4[m
[32m+[m[32mgoogle-auth-oauthlib #==1.1.0[m
[32m+[m[32mgoogle-auth-httplib2 #==0.1.1[m
[32m+[m[32mgoogle-api-python-client #==2.108.0[m
[32m+[m[32mtabulate #==0.9.0[m
