# Macro

Okay so this excel file named Thike.xlsx is a macro enabled file which is why its saved as a .xlsm

The macro was created by using VBA. All you have to do is click on the Upload to S3 Button which will enable the macro. 

The end result will be the uploading of the green cells to a S3 bucket on AWS as a json file.

The macro itself, however, does not do any of this. All it does is enables a python script named "bestest.py" 

The python script can read the contents of the cells of the excel sheet, select the data specified from the cells we want, then convert it into 
a JSON file. The JSON file is named as  <today’s date>_<tab name>.json. After that, it uploads the json file automatically to the S3 bucket.
It will not save the JSON file to the sytstem though because there is no need for it, however a couple of tweaks will allow it.

Now the python script works on the belief that you have your python exectuable path setup on your system as a part of the environmental variables. Additionally, pip install of boto3 and openpyxl is required as well. 

Within the python script itself, we must provide the path to our excel file. Change it according to the requirements. Additionally, the S3 bucket name must be provided as well so that the system knows where to upload the file to. This means that the S3 bucket has been created beforehand, so that its path can be pasted into the python script. 

The variable s3_key basically names the file as its going to appear on the S3 bucket. So in this case, <today’s date>_<tab name>.json. If however your name it as "cool.json" then the object(json file) will appear as such on your S3 bucket. 

Now, lets talk about the VBA script. The scirpt itself is very simple and it connects to the python executable and python script that is already on your system. So change the path of the python executable(python.exe) and paste it and then do the same for the python script(bestest.py) path as well. Most likely your python scirpt will be found in the downloads folder. If you cannot find the python.exe then simply open cmd and type "where python" and the path will show up from where you can copy said path.




This of course will only work if you have an AWS account setup already. [Here]([url](https://www.youtube.com/watch?v=XhW17g73fvY)) is a good video. 

After setting up your AWS account you must create an IAM user, create a user, give the user permissions to interact and create s3 buckets and then create an actual S3 bucket. 

Afterwards, we must install and setup AWS CLI. Then configure it so that the system can upload files as objects to your S3 bucket. 

And, ideally this should upload the json file to your intended destination.

