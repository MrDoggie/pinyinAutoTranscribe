# pinyinAutoTranscribe
Taking regular Pinyin as input, transcribe into modified Pinyin and Duanmu Pinyin.

To use the tool, make sure you have a copy of the original file.

*Be extreamly careful that your original file is not called data.xlsx*

Then create a new excel file called data.xlsx

The name and extension name should be exactly the same as above. Otherwise the code will not work.

Put all regular pinyin input such as 'zhuan' 'yi' into the first column. One word per row.

Notice: the input should have no special notation including space. Otherwise the code will fail to transcribe.

Then run the code.

The second column will be the onset.

The third column will be the modified pinyin glide.

The fourth column will be the modified pinyin rhyme.

The fifth column will be the Duanmu pinyin glide.

The sixth column will be the Duanmu pinyin rhyme.

error matching will be shown as ###. Go back and make sure your input is correct. If you find your input is correct and the program is doing unexpected things, email wangjunyi@ucla.edu

Finally manually copy the result back to your original file. Hopefully this will save you a ton of time and largely improve data's correctness.

This code is originally designed for Brice Roberts' Mandarin Pop Rhyme Project @UCLA


===================================
Aug 25th 
* Add -v matches. Obviously I forget the cases such as lv and nv. Now I've added them into the file. 
* Now you can customize your input file name, your input column number and your output column number easily
