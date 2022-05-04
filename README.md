# Automating_Without_API
This code is used to download daily csv files from websites that I don't have API access to. While the code is specific to my directories and keyring data, you can still find utility in this code by looking at the usage of explicit waits, the try and except portion of the download-report function (used to click the extra button for zip files only if the download is in a zip file), and the idea and functionality behind using the holding directory (holding_dir) to explicitly wait for the csv to complete downloading before moving on to the next file. Looking at how the script accounts for the possibility of files being downloaded in zip files and conditionally unzips them through the move function may also be beneficial.
