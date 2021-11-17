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

## Addying the python script to the context menu.
1. Get [EcMenu](https://www.sordum.org/7615/easy-context-menu-v1-6/)