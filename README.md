# hwp-Automator
Automator tool for hwp files. Developed for hagwons
The program uses API provided by hancom to access and modify hwp files.

### What for?
This is used to automate the process of tedious and time consuming task of test paper generation. One test paper takes 100 senteces from text book. There are 20 questions with 5 of theses sentences. One sentece got its keyword replaced with other keywords in the textbook, which should not make any sense. This sentence is the answer students identify.

### Current Feature:
Copy **test.hwp** file as a format and make a new hwp file called **test_out.hwp**
Takes a 100 randomly selected non-repeating senteces from **source.xlsx** file and put those sentences into **test_out.hwp**. 

### Next:
1. Underline corresponding keywords on **test_out.hwp**.
2. Randomly assign answers and replaced words. i.e. make answer key on the back of the test paper.
3. Make new file __new_source.xlsx__ that does not contain used sentences.

### Why such inefficient codes?
There are two answer for this. 
1. It's my first time using HwpObject and some efficient algorithm I wanted to use did not work.
2. I was in very hurry because the program needed to be used right away to save time. I could not watch my mother and sister coming home at 3 in the morning after making test papers. 

