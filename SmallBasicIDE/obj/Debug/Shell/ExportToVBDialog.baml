   M S B A M L   `   `   `  ����=  9SB, Version=0.1.0.0, Culture=neutral, PublicKeyToken=nullK*clr-namespace:Microsoft.SmallBasic.UtilityMicrosoft.SmallBasic.Utility  5   	   G(clr-namespace:Microsoft.SmallBasic.ShellMicrosoft.SmallBasic.Shell  5
   	   1    +Microsoft.SmallBasic.Shell.ExportToVBDialog   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   9lu*clr-namespace:Microsoft.SmallBasic.Utility     5   	   :shell(clr-namespace:Microsoft.SmallBasic.Shell     5
   	   -   5          �Name$  exportToVBDialog��5   	    � WindowStartupLocation$ CenterScreen=�5   	    � WindowStyle$ SingleBorderWindow=�5   	    � 
ResizeMode$ NoResize=�5   	    � SizeToContent$ WidthAndHeight=�6    � ShowInTaskbar .   � Title 5   	   5   /Microsoft.SmallBasic.Utility.LocalizedExtension  *
Graduate+.��5      ��  �� Orientation$ Vertical=�6   ���	#FFFFFFFF$	��400��6?   .+�5   
   � $	��180��6    � ColumnDefinitions 5      �� 5      $
��Auto �6#   �� 5      $��* �6#   .+�5      �� -   5        $��iconImageControl��5      $?��9pack://application:,,/SB;component/Resources/VBExport.png��6,   $��0��6o   $��64��6   $��64��6�   $	��Top=�6�   $��8q�6�   �� 5      $ Vertical=�6   $��1��60   $	��Top=�6@   .+�5      �� $��2,10,2,2q�6   $��WrapWithOverflow=�6.   $��Segoe UI)�6N   $��16(�6d   ���#002ab5��6�     *
ExportToVB+�� 5      -   5        $��optionalContentControl��5      $��2q�6:   $��Segoe UI)�6E   $��12(�6[   $
��Wrap=�6i   ��6}     *ExportToVBDescription+�� 5      $��2,8,2,2q�6   ��6-     *OutputLocation+�� 5      $ 
Horizontal=�6   .+�5      �� -   5        $��locationTextBox��5      $	��280��65   $��2,2,0,2q�6A   	 �� 
IsReadOnly	 . 6R   �� 5      -   5        $��20��5      $��22��6(   .��6J   ...6O   5      �� 5      -   5        $��statusPanel��5      $��0,12,4,0q�60   $ Vertical=�6B   $y�	Collapsed=�6Y   .+�5      �� -   5        $��
statusText��5       $��2q�62   � 5      $��2,2,8,2q�6"   $��8��63   
 � IsIndeterminate
 . 6>   5      �� 5       -   5        $��
errorPanel��5       $��0,12,4,0q�6/   $ Vertical=�6A   $y�	Collapsed=�6X   ���	#FFFF0000.+�5!      �� $��2q�6    $
��Wrap=�6+   ��6?     *FoundErrorsInProgram+5"      �� 5#      -	   5        $��notInstalledPanel��5#      $y�	Collapsed=�66   $��0,12,4,0q�6M   ���	#FFFF0000.+�5$      �� $
��Wrap=�6    ��64     *
VBNotFound+�� 5%      -
   5        .�5%   /   &$http://www.microsoft.com/express/vb/6U   5&      5'      5(      5)      �� 5*   
   $��Stretch=�6   $��
0, 8, 0, 0q�6/    �� Padding$ 4q�6C   ���sc#0.7, 0.6, 0.6, 0.6$��0,1,0,0q�6s   ���sc#0.7, 0.83, 0.83, 0.83._�5+      �� $ 
Horizontal=�6   $��Right=�62   .+�5,      �� -   5        $��continueButton��5,      $��80��6/   $��4,2,0,2q�6:    �� 	IsDefault . {�.  ��6n     *ButtonContinue+5-      �� 5.      -   5        $��cancelButton��5.      $��80��6-   $��0,2,4,2q�68    �� IsCancel . ��6Y     *ButtonCancel+5/      50      51      52      