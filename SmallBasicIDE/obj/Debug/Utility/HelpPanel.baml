   M S B A M L   `   `   `  ����=  9SB, Version=0.1.0.0, Culture=neutral, PublicKeyToken=null;"clr-namespace:Microsoft.SmallBasicMicrosoft.SmallBasic  5   �   K*clr-namespace:Microsoft.SmallBasic.UtilityMicrosoft.SmallBasic.Utility  6�   ,    &Microsoft.SmallBasic.Utility.HelpPanel   6   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    68   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     6z   1sb"clr-namespace:Microsoft.SmallBasic     6�   9lu*clr-namespace:Microsoft.SmallBasic.Utility     6�   -   5          �Name$  	helpPanel��5      � 	Resources 5         b2v  paramsNameStyle �� 
TargetType*   $Microsoft.SmallBasic.Utility.TextRun �� Property �� Value  paramsDescriptionStyle  membersStyle  
usageStyle  	typeStyle  codeExampleStyle�� 5      %  &	        &	      &	 �     &	     &	 �    &	 �    &	     �� 5      6%   �� 5      "  6&   .�5   
   ��  � ��$ 18(�6%   �� 5   
    � �� �#346D846'   �� 5   
    � ��$
 Bold"�6'   5	   	   �� 5
      "  6-   .�5   
   ��  � ��$ 16(�6%   �� 5   
    � �� �	#FF0000006'   5   	   �� 5      "  6#   .�5   
   ��  � ��$ 16(�6%   �� 5   
    � �� �#5555556'   �� 5   
    � ��$
 Bold"�6'   �� 5   
    �@��BaselineAlignment$	 Top=�65   5   	   �� 5      "  6!   .�5   
   ��  � ��$ 16(�6%   �� 5   
    � �� �#d845236'   �� 5   
    � ��$
 Bold"�6'   5   	   �� 5      "  6    .�5   
   ��  � ��$ 22(�6%   �� 5   
    � �� �#346D846'   �� 5   
    � ��$
 Bold"�6'   5   	   �� 5      "  6'   .�5   
   ��  � ��$ 15(�6%   �� 5    
    � �� �	#FF0000006'   �� 5!   
    � ��$ Consolas)�6'   5"   	   5#      .+�5%      �� ���	#80dfefff �� CornerRadius$ 1��6"   ���	#d0dfefff$��1q�6K   $��2��6_   ._�5&      ,� $��0q�6   .��5'      /� -   5         �Name$ helpDocument��5'       /� 
FontFamily$ Segoe UI)�5(       /� FontSize$ 16(�6+   	 /� PagePadding$	 4q�5)      
 /� TextAlignment$

 Left=�6%    /� IsHyphenationEnabled .    Microsoft.SmallBasic.App   FlowDirection /� FlowDirection# Z 5+      .T�5,   
   J�  �� Margin$ 2q�6   .*�6   .*�5-      �� -   5        $��symbolImage��5-      $��64��6%   $��64��60      5.      -   5          Name$ 
symbolType��5.      $��10(�6-    �� BaselineAlignment$	 Top=�5/      �� 50        51      $��20(�6   ���#346D84   Text 52      ��  �� Path$ Display��� 53        54      $
��Bold"�6   ���#555555 55      �� $ Documentation.Summary� �� 56      $��Stretch=�6   $��0,4,0,4q�66   $	��800��57      58      J� 59   
   -   5        $ detailsParagraph��59      5:   	   5;      5<      5=      