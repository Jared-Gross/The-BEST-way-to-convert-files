# The-BEST-way-to-convert-files
The best way to convert files on your computer, be it .pdf to .png, .pdf to .docx, .png to .ico, or anything you can imagine.

# Goal?
Are you sick and tired of going to an online website to convert certain documents? 

The tedious process, and the length you have to go through has driven me up the wall.

I made a [gist](https://gist.github.com/JareBear12418/d4094a87b17364952937f54ba1728a0a) a while ago explaining how you can do this through the CLI, which don't get me wrong is a great way to do it if you have a terminal on hand. But I found an even better way to do this.

I will be using the same script/converter in the gist mentioned above just for simplicity sake, the same applies to ANY file conversion you can think of.

# How?
Context Menus.

**NOTE** This is WINDOWS SPECIFIC!

How can we add to the context menu?
The answer is in [EcMenu](https://www.sordum.org/7615/easy-context-menu-v1-6/). It's free, and it works super well!

Let's go over how to set up a python script and run that through the Context Menu.

# Setup
Any conversion script has these requirements:

## What do we need?
1. Convert file from something TO something else.
    - This doesn't need to be spesfically `.png` to `.ico`. It can be a fairly complex conversion, say `.pdf` to images.
2. Input file
    - We need to get the input file we want to convert, the rest is handled by the script.
3. Current directory
    - We need the directory of the file you want to convert. This way the program works no matter where the file is located. Alternativly you can set a CONSTANT output directory if you so please in the script itself. For this example we will be working in the directory of where the file is located itself.

## Setting up python script
**main.py**
```python
# Required libraries
import sys, os

def convert(file_name: str, output_dir: str):
    ...

# Get the file name
file_name: str = sys.argv[-1].split('\\')[-1]

# Get directory of where the file is located 
directory_of_file: str = os.getcwd()

# Start the conversion process
convert(file_name, directory_of_file)
```

How do we run this? Why are we not using `argparse`?

To run this script we do:
`python main.py this_is_my_file.txt`

We don't use `areparse` because we can't directly add arguments to the context menu. `argparse` we need a `-f this_is_my_file.txt` to specify the file, but we don't need that, we know we input only one file, so we can just use `sys`.

`sys.argv[-1].split('\\')[-1]` This gives us the file that you right clicked on.

We then need compile the python script using:

`pyinstaller -F pdftoword.py`

In this script i'm using the terminal for feedback, you can use what ever you want, you can use `tkinter`, `qt`, or not give anyfeed back at all, to do that simply use:

`pyinstaller -F --noconsole pdftoword.py`

The `.exe` will be in a `dist` directory where you the python script is located.

## Addying the python script to the context menu.
1. Get [EcMenu](https://www.sordum.org/7615/easy-context-menu-v1-6)
2. Go to list Editor

![image](https://user-images.githubusercontent.com/25397800/142265460-6c3047bc-a389-4d51-b743-9821bc982953.png)

3. Scroll down to **File Context Menu**

![image](https://user-images.githubusercontent.com/25397800/142265628-ef5ff283-9b47-42af-b3da-a99fba2ef0c2.png)

4. Make sure you select **File Context Menu** and press **Add New** and Browse to where the `.exe` file is located.

![image](https://user-images.githubusercontent.com/25397800/142265847-061e1d4d-b885-4dca-8f23-684304bd62f1.png)

5. Press **Save Changes**

![image](https://user-images.githubusercontent.com/25397800/142266216-52010316-fcbb-4141-a002-73130064e18d.png)

6. Make sure it's checked!

![image](https://user-images.githubusercontent.com/25397800/142266363-5ce4fe62-35e2-41d1-a4f4-86d19a2e2477.png)

7. Press **Apply Changes**

![image](https://user-images.githubusercontent.com/25397800/142266458-9942f641-095e-4dc3-b4ce-0a5146fd8d66.png)

And your done. Right click on any file and you should see it!

# Summary
This is just a 'quality of life' solution. You can go above and beyond with this method of converting files, you can add a whole suite of file conversions if you so want too.

