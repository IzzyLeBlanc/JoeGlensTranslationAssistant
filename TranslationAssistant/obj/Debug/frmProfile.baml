   M S B A M L   `   `   `  ����  TranslationAssistant%    TranslationAssistant.frmProfile   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5        � Title$  Profile Setting��5      $��501.253��6   $��345.432��6.    � WindowStyle$ 
ToolWindow=�6>   $��Microsoft YaHei)�5      $
��11.5(�5   	    � 
ResizeMode$ NoResize=�6    � 	Resources 5         CenterGridHeaderStyle �� 
TargetType=  7System.Windows.Controls.Primitives.DataGridColumnHeader �� Property �� Value��%V   &	        �� 5   
   "  6.   .�5	      ��  � ��$ Center=�6;   5
      5      .��5      � .+�5   
   �� $��Parse Characters��6   $
��Left=�6+   $��29��6F   $��10,8,0,0q�6R   $	��Top=�6d   $	��111��6|   �� 5   
   -   5         ��Name$ parseCharactersInput��5      $
��Left=�60   $��25��6K   $��
126,12,0,0q�6W    �� TextWrapping$
 Wrap=�6k   $	��Top=�6   $	��166��6�   $��、 。 ？ ！ 「 」��6�   &   System.Windows.Media.TextOptions	   TextFormattingMode$	 Display=�6�   �� 5   
   $��Profile��6   $
��Left=�6"   $��23��6=   $��	10,44,0,0q�6I   $	��Top=�6\   $	��101��6t   �� 5   
   -   5        
 ��Name$
 	comboList��5      $
��Left=�6&   $��25��6A   $��
126,44,0,0q�6M   $	��Top=�6a   $	��166��6y   �� 5   
   -   5         ��Name$ 	btnCreate��5      $��
Create New��6$   $
��Left=�69   $��24��6T   $��	69,88,0,0q�6`   $	��Top=�6s   $��80��6�   �� 5   
   -   5        $ 	btnDelete��5      $��Delete Profile��6$   $
��Left=�6=   $��24��6X   $��
179,88,0,0q�6d   $	��Top=�6x   $��80��6�   �� 5   
   -   5        $ btnSave��5      $
��Save��6"   $
��Left=�61   $��126,430,0,0q�6L   $��80��6a    �� 	IsDefault . $��19��6}   $	��Top=�6�   &   System.Windows.Controls.DataGrid  5   
   -   5          Name$ 
dgvPhrases��5      $
��Left=�6'   $	��298��6B   $��
10,127,0,0q�6O   $	��Top=�6c   $	��304��6{    J� IsManipulationEnabled .    AutoGenerateColumns .     Columns 5      0  *System.Windows.Controls.DataGridTextColumn  5      ,  &System.Windows.Controls.DataGridColumn   Header$ Phrases��6@      Width5  /System.Windows.Controls.DataGridLengthConverter$
 0.5* 6Q   1  +System.Windows.Controls.DataGridBoundColumn   Binding 6%   �� *Phrase+6Q     5      $ Translation��6E   $
 0.5* 6Z    6%   �� *Translation+6Z   5      5      5      