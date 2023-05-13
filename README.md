# hwp-Automator
Automator tool for hwp files. Developed for hagwons
The program uses API provided by hancom to access and modify hwp files.

### What for?
This is used to automate the process of tedious and time consuming task of test paper generation. One test paper takes 100 senteces from text book. There are 20 questions with 5 of theses sentences. One sentece got its keyword replaced with other keywords in the textbook, which should not make any sense. This sentence is the answer students identify.

### Current Feature:
1. Copy **test.hwp** file as a format and make a new hwp file called **test_out.hwp**
Takes a 100 randomly selected non-repeating senteces from **source.xlsx** file and put those sentences into **test_out.hwp**. 
2. Iteratively Underline corresponding keywords on **test_out.hwp**.
3. Randomly assign answers and replaced words. i.e. make answer key on the back of the test paper.

### Next:
1. Create GUI for the program.
