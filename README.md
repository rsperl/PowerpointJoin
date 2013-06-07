## About
I am grateful that the Lord Jesus Christ has given me the ability to write something useful that helps others. He has also given me a job that I enjoy, so I am happy to make this available for free.

Should you find this application useful and want to give back, I humbly ask that you consider making a donation to one of the following organizations:

*  [Christian Counseling Education Foundation](http://www.ccef.org/donate): CCEF's mission is to "Restore Christ to counseling and counseling to the church".

* [White Horse Inn](http://www.whitehorseinn.org/partnerships/support-us.html): The White Horse Inn is a blog and podcast that aims teach Christians to "know what they believe and why they believe it".

## GUI
At some point I will try to build a gui for easier use.

## Use case
Given a several Powerpoint presentations, you want to merge them together. You could do it within Powerpoint, but perhaps you only want slides 1, 3, 8-9, and 20 from one deck, all of the next deck, and a few more onesy-twosy slides from various other decks.

## Requirements
This script is written in Perl. Perl is available for free from either [ActiveState](http://www.activestate.com/activeperl/downloads) or [Strawberry](http://www.strawberry.com). Once that is installed, you will also need to install two modules that may not already be installed. From a command prompt, type

    perl -MCPAN -e "install Win32::PowerPoint"

Once that completes, double check that the second module is installed by typing

    perl -MCPAN -e "install Win32::OLE"

Now you're ready.

## Configuration file
The configuration file defines which the source charts and which slides from those files are used. Blank lines and lines that start with # are ignored.


**Example: insert all slides from each file**

    file=files\a.pptx
    file=files\b.pptx
    file=files\c.pptx

**Example**

    file=files\a.pptx
    slides=1-2

    file=files\b.pptx
    slides=3

    file=files\c.pptx
    slides=1,3
    
## Use
After downloading the zip file, extract them to your hard drive, for example, in c:\PowerpointJoin. Open a command prompt (Start | All Programs | Accessories | Command Prompt). Go to this directory by typing 

    cd c:\PowerpointJoin

Make sure that everything works by typing

    PowerpointJoin.pl

The response should look like the following:

    C:\PowerpointJoin>PowerPointJoin.pl
    Usage: C:\PowerpointJoin\PowerPointJoin.pl --conf
    ig <config_file> --output <output.ppt>


    The config file should have the following format:

    file=filename1.ppt
    slides=1-4,6-8

    file=filename2.ppt

    file=filename3.ppt
    slides=3

    The start and end lines are optional. start defaults to 1 and end
    defaults to the last slide in the file. The blank line betwen
    sections is option, but makes it readable if you have a lot of
    start/end lines.

To process your config file, assuming your config file is named "config.txt", type

    PowerpointJion.pl --conf config.txt --output merged_charts.pptx

The merged charts will be in the file "merged_charts.pptx"

