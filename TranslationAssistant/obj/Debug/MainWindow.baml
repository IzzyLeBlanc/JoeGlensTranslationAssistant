   M S B A M L   `   `   `  ����O  KTranslationAssistant, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null;"clr-namespace:TranslationAssistantTranslationAssistant  5   	   %    TranslationAssistant.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   4local"clr-namespace:TranslationAssistant     5   	     � Icon$-  'pack://application:,,,/Resources/TA.ico��5   	    � Title$  Translation Assistant Tool��5	   	   $	��686��5
   	   $	��566��5   	   $��Microsoft YaHei)�5   	   $
��12.5(�5   	    � WindowStyle$ ThreeDBorderWindow=�5   	   $	��500��5   	   $	��564��5   	   .��5      ]� ��.  .+�5      o� -   5         o�Name$
 menu��5      $��23��6    o� 
IsMainMenu . $	��top=�6<   .@�5      m� $
��File��6   .@�5      m� -   5         m�Name$ menuNew��5      $��New (CTRL+N)��6,   m� 5      -   5        $ menuOpen��5      $��Open (CTRL+O)��6-   m� 5      -   5        $ menuSave��5      $��Save (CTRL+S)��6-   {�.  6D   5      m� 5      $��Settings��6   .@�5      m� -   5        $ menuProfile��5      $��Profile (CTRL+P)��60   m� 5      -   5        $ 
menuPhrase��5      $��Phrase (CTRL+L)��6/   m� 5      -   5        $ menuProgress��5      $��Show Progress��61   m� 5      -   5        $ 	menuOnTop��5      $��Always On Top��6.   m� 5      -	   5        $ menuTTS��5      $��Text-To-Speech��6,   .@�5      m� -
   5        $ 	menuTTSJP��5      $��Japanese��62   {�.  6D   m� 5      -   5        $ 	menuTTSCN��5      $��Chinese��62   {�.  6C   5       5!      m� 5"      $��Special Punctuations��6   .@�5#      m� -   5        $ 
menuSingle��5#      $$��Single Qoute : 「　」  (F1)��6/   m� 5$      -   5        $ 
menuDouble��5$      $$��Double Qoute : 『　』  (F2)��6/   m� 5%      -   5        $ menuLenticular��5%      $"��Lenticular : 【　】  (F3)��63   m� 5&      -   5        $ menuEllipsis��5&      $��Ellipsis : …  (F4)��61   m� 5'      -   5        $ menuWave��5'      $��Wave Dash : 〜  (F5)��6-   m� 5(      -   5        $ 
menuSTitle��5(      $*��$Single Title Bracket : 〈 〉  (F6)��6/   m� 5)      -   5        $ 
menuDTitle��5)      $)��#Double Title Bracket : 《 》 (F7)��6/   m� 5*      -   5        $ menuDash��5*      $��Long Dash : ー (F8)��6-   5+      m� 5,      -   5        $ menuClipboard��5,      $��Clipboard (CTRL+I)��6.   {�.  6J   m� 5-      -   5        $ 	menuAbout��5-      $��About��6*   5.      �� 5/   
   -   5         ��Name$ 	statusBar��5/      $��23��6'   $��Bottom=�63   $��Bottom=�6K   .@�50      �� -   5         ��Name$ completionStatus��50      $
��Left=�66   $��0% Complete��6Q   �� 51      -   5        $ 
LineStatus��51      $
��Left=�60   $��Line: xxxx/xxxx��6K   �� 52      -   5        $ wordCountStatus��52      $
��Left=�65   $��
xxxx Words��6P   �� 53      -   5        $ 	filesaved��53      $��Right=�6/   $��
File Saved��6K   54      �� 55   
   -   5         ��Name$ 	reviewTop��55      	 �� TextWrapping$
	 Wrap=�6&   $����HOW TO USE:
Creating New File for Translations:
1.)Click File->New
2.)Copy Raw from Source and paste it on the textbox
3.)Click Create
4.)Save File into the desired location
5.)Translate
6.)Once done translating click on Clipboard. Paste on your preferred editor or your blog for final edit and posting

Creating New Profile
1.)Click Settings->Profile
2.)Click New Profile
3.)Type profile name
4.)Click OK

Adding Phrases to Profile:
Option 1:
1.)Click Settings->Add Phrase
2.)Enter Raw on Phrase textbox
3.)Enter translation of phrase in Translation textbox
4.)Click Save
Option 2:
1.)Click Settings->Profile
2.)Add Raw text and Translated Test on their respective boxes at the end of table
3.)Click Save

Deleting Phrases
1.)Click Settings->Profile
2.)Double-Click phrase you want to delete
3.)Confirm Delete
4.)Click Save

Editing Phrases
1.)Click Settings->Profile
2.)Click on a phrase, and click once again (make sure to have a pause between clicks or else it will be considered a double click and will ask to delete an entry) 
3.)Modify entry
4.)Click Save

Use Spelling Suggestions
1.)Right-Click the red underlined word
2.)Choose the correct spelling you meant to type from the context menu

Adding to Custom Dictionary
1.)Highlight word to add to custom dictionary (Double click on the word)
2.)Press CTRL+J and confirm the action or Right-Click and Add To Dictionary

Special Punctuations
1.) You can click on the Special Punctuations menu to insert the desired punctation or press the corresponding Function key (F1 - F5)��6:   $	��Top=�6  $��5q�6&  $	��320��61  
 �� 
IsReadOnly
 . " �� VerticalScrollBarVisibility$ Visible=�6P  �� 56   
   -   5        $ currentRawLine��56      $��52��6+   $
	 Wrap=�67   $��5q�6K   $	��Top=�6V   
 . $ Visible=�6}    �� IsUndoEnabled .  6�   ]� 57   
   $	��Top=�6    +� IsFocusScope . .+�58      �� -   5        $ currentTranslatedLine��58      $��52��66   $
	 Wrap=�6B   $��5q�6V   $	��Top=�6a    �� 	IsEnabled . $ Visible=�6�   .��6�    � ContextMenu 59      �� 5:      .@�5;      m� .@�6#   dummy6*   5<      .��5=   '   5>      �� 5?      -   5        $ reviewBottom��5?      $	��137��6,   $
	 Wrap=�6:   $��5q�6N   $	��Top=�6Y   
 . $ Visible=�6�   $����NAVIGATION CONTROLS:
ENTER Key or PgDn Key  = Move down next sentence.
PgUp Key = Move up to previous sentence
CTRL+Left Key = Highlight previous parsed phrase in a sentence, if this was the first parsed phrase it will revert the sentence to original raw if there was any glossary swapping done
CTRL+Right Key = Highlight next parsed phrase in a sentence, if this was the original raw that was reverted from a glossary swap it will swap any glossary terms
CTRL+Home Key = Go to the the first line
CTRL+End Key = Go to the most recent un-translated sentence��6�   5@      5A      5B      