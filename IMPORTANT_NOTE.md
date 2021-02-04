
<p align="center"><img width="150" height="160" src="https://www.drupal.org/files/project-images/Google-API.jpg"/></p>

<br><br>

THIS IS REQUIRED TO MAKE THE PYINSTALLER .exe WORKS!!

Since PyInstaller .spec file requires an explicit link for the Google library files in order to generate a working .exe you are supposed to generate one yourself on your machine. 
the .spec file, apart from any other spec you will eventually add, should have added in the "datas" key's array a tuple composed like this ( 'absolute_path_to_google_api_python_client-x.y.z.dist-info', 'google_api_python_client-x.y.z.dist-info').

DO NOT use the .spec file in the repository to build your BusinessCat.exe, either use it as a template to manually replace all absolute paths with yours or just ignore it! 

