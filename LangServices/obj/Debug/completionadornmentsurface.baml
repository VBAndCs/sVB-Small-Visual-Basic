   M S B A M L   `   `   `  ����O  Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem  5   �   J FLanguageService, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null[2clr-namespace:Microsoft.SmallBasic.LanguageService$Microsoft.SmallBasic.LanguageService 6�   E   ?Microsoft.SmallBasic.LanguageService.CompletionAdornmentSurface   6	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    6S   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     6�   6sys&clr-namespace:System;assembly=mscorlib     6�   @l2clr-namespace:Microsoft.SmallBasic.LanguageService     6�   -   5          ��Name$  completionSurface��5   <   � 	Resources 5       �� DataType@  :Microsoft.SmallBasic.LanguageService.CompletionItemWrapper �� Orientation �� Path �� Triggers �� Value �� Binding �� 
TargetName	 �� Property
 �� Value �� 
TargetType �� 
TargetType �� CornerRadius �� Triggers P� Property P� Value �� 
StartPoint �� EndPoint � Offset � Color�� 5      %�	  (��       5   P   *, +)'�� �    �� 6   "  6P   .R�5   
   �� $ 
Horizontal=�6   .+�5      �� $��	itemImage��6   $G��Apack://application:,,/SB;component/Resources/IntellisenseItem.png��6#   $��14��6n   $��14��6y   $��1,2,1,1q�6�   �� 5      $��Segoe UI)�6   $��11(�6,   $��2,1,0,0q�6:   $��CharacterEllipsis=�6K   $	��100��6l   ��6{   �� $
 Name�5	       5
   
   �� 5      $ Event��6    6&   �� $ 
SymbolType�.a�5      �� $ 	itemImage��6   	 � ��$H
 Bpack://application:,,/SB;component/Resources/IntellisenseEvent.png��6D   5      �� 5      $ Keyword��6    6(   �� $ 
SymbolType�.a�5      �� $ 	itemImage��6   	 � ��$J
 Dpack://application:,,/SB;component/Resources/IntellisenseKeyword.png��6D   5      �� 5      $ Method��6    6'   �� $ 
SymbolType�.a�5      �� $ 	itemImage��6   	 � ��$I
 Cpack://application:,,/SB;component/Resources/IntellisenseMethod.png��6D   5      �� 5      $ Property��6    6)   �� $ 
SymbolType�.a�5      �� $ 	itemImage��6   	 � ��$K
 Epack://application:,,/SB;component/Resources/IntellisenseProperty.png��6D   5      �� 5      $ 
Subroutine��6    6+   �� $ 
SymbolType�.a�5      �� $ 	itemImage��6   	 � ��$M
 Gpack://application:,,/SB;component/Resources/IntellisenseSubroutine.png��6D   5      �� 5      $ Variable��6    6)   �� $ 
SymbolType�.a�5      �� $ 	itemImage��6   	 � ��$K
 Epack://application:,,/SB;component/Resources/IntellisenseVariable.png��6D   5      �� 5      $ Label��6    6&   �� $ 
SymbolType�.a�5      �� $ 	itemImage��6   	 � ��$H
 Bpack://application:,,/SB;component/Resources/IntellisenseLabel.png��6D   5      5!   	   �� 5"      " ��6+   .�5#   
   �� 	 �@�FocusVisualStyle
 6>   Z� �� 5$   
   	 � ��
 5%      �� 5&      " ��6   .R�5'      �� $��Border��6   $��1q�6%   $ 2��69   ���	#00FFFFFF���	#00FFFFFF._�5(      �� 5)       5*      P� 5+       �@��
IsSelected . . �5,      �� 	 � ��$

 Bold"�6;   �� 5-      $ Border��6   	 � ��
 �#A8D8EB6M   �� 5.      $ Border��6   	 � ��
 5/      �� 50      $	 0,0&�6,   $	 0,1&�6=   .��51      � $ 0Q�6'   $ #E7EFF3��62   � 52      $ 1Q�6'   $ #DAEBF3��62   53      55      56      58      5:      5;   	   5<      .+�5?      � -   5        $��CompletionPopup��5?       � 	StaysOpen .  $	��360��65   $	��140��6A   $��Relative=�6N    � AllowsTransparency . |�.  $
��Fade=�6�   .��5@      �� -   5        $��popupContent��5@      $��1q�6!   ���#999999$ 5��6K   ��5A      �� 5B   
   $	 0,0&�6   $	 0,1&�6/   .��5C      � $ #A9D5E7��6   $ 0Q�6)   � 5D      $ #CFE9F3��6   $	 0.4Q�6)   � 5E      $ #D1EAF3��6   $ 1Q�6)   5F      ._�5H      � $	��360��6   ��5I   
   L� 5J       L� Drawing 5K      I� 5L      ��5N      �  � Geometry� ��hS��.�	��q�.�	���&�~p@�.�	�,|�.�	ӐI
�.�	8��.�	X���1	�xTw`'��A�  �l�
�x�(�
��a��	 �-��@ ��|Vi�t�,L����@#�@B  ��Ӏ� ����l ����@ܵ���	 (��
@�el�
�����A� Q%xTw�IX���Ŕ �8�    �I
�,|�&��    Қ !8�(��    ���ا ���ا ����RW�C��\����0:�Ӏ���.�	���.�	hS��.�	
 � Pen 5O      >� 5P       >� LineJoin$ Round=�6    >� Brush �#A8D8EB >� 	Thickness$ 1Q�6=    � Brush 5R      �� 5S      $	 0,0&�6,   $	 0,1&�6=   .��5T      � $ 0Q�6'   $ #E7EFF3��62   � 5U      $ 1Q�6'   $ #DAEBF3��62   5V      5X      5Z      5\      .+�5_   
   ]� .+�5`      �� $��Right=�6   ._�5a      �� $	��138��6   $��0,0,2,0q�6)   ��6:   �� $  SelectedItem.Documentation� �� ElementName$ CompletionListBox����5b      �� 5c      F  @Microsoft.SmallBasic.LanguageService.CompletionItemDocumentation"  6   .R�5d      �� $
��Auto=�6!   $��2q�6D   .��5e      �� $��Stretch=�6!   ��6?   �� $ TextFlowDirection�$ completionSurface��.+�5f      �� $
��Bold"�6"   $��20(�64   $��0,0,2,0q�5g   "   $��CharacterEllipsis=�63   ���#346D84��5i   "   �� $ SelectedItem.Name�$ CompletionListBox���� 5j      ���#e0e0e0$��0,0,0,1q�65   $��Stretch=�6O   �� 5k      $��0,4,2,0q�6"   $��16(�63   $
��Wrap=�5l   "   ��5m   "   �� $ Summary�5n      5o      5p      5r      5s      7  1Microsoft.SmallBasic.LanguageService.CircularList  5t      -   5          Name$ CompletionListBox��5t      7  1Microsoft.SmallBasic.LanguageService.CircularList    SelectionChanged$$  OnCurrentCompletionItemChanged��5u      !   MouseDoubleClick$#! OnCompletionListDoubleClicked��5v      |�.  $��0q�5w   "   ���	#00FFFFFF-  'System.Windows.Input.KeyboardNavigation"   DirectionalNavigation$" 	Contained=�5y      $
��Left=�5z      $	��220��6&   5{      5|   	   5}      5~      5      