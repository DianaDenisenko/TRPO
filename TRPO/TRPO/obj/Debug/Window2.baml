   M S B A M L   `   `   `  ����?  ;TRPO, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:TRPOTRPO  5   	       TRPO.Window2   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   b ^MaterialDesignThemes.Wpf, Version=4.0.0.2422, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   
   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   $localclr-namespace:TRPO     5   	   -   5          �Name$  	WindowReg��5   	    � Title$ Window2��5
   	   $��516.568��6   $��459.171��6*    � 
ResizeMode$ NoResize=�6:    � WindowStartupLocation$ CenterScreen=�6P    � WindowStyle$
 None=�6u   ��5      �� 5   
    �� EndPoint$ 0.5,1&�6    �� 
StartPoint$ 0.5,0&�6/   .��5      �  � Color$ 	#FF511D1D��6    � Offset$ 0Q�6-   � 5      $ 	#FF954747��6   $ 1Q�6-   5      .��5      � $��	0,0,583,0q�6   .+�5   
   � 	 J� RenderTransformOrigin$	 0.5,0.5&�6   $��-2,1,-459,-15q�6/   ��5      �� 5      $ 0.5,1&�6&   $ 0.5,0&�67   .��5      � $ 	#FF1B0C0E��6#   $ 0Q�65   � 5      $ 	#FF6C2E36��6#   $ 1Q�65   5      z�5      V� 5      .~�5      �� �� 5      �� 5      
 �� Angle$
 -0.117Q�6&   T� 5      5      .+�5!      �� -   5         ��Name$	 TB1��5!      $
��Left=�6#   .  (MaterialDesignThemes.Wpf.TextFieldAssist   UnderlineBrush
 �����)  #MaterialDesignThemes.Wpf.HintAssist   Hint$ Имя��6t   $��Times New Roman)�6�   $��41��6�   $��125,184,0,0q�6�    �� TextWrapping$
 Wrap=�6�   $	��Top=�6�   $	��230��6  ���	#FFE0E6E7$��25(�6=   �� 	MaxLength$ 15��6e  ��6  Z� 6t  �� 5"      -   5        $	 TB2��5"      $
��Left=�6#   
 �����$ Фамилия��6t   $��Times New Roman)�6�   $��41��6�   $��125,240,0,0q�6�   $
 Wrap=�6�   $	��Top=�6�   $	��230��6	  ���	#FFE0E6E7$��25(�6B  $ 15��6j  ��6  Z� 6y  �� 5#      -   5         ��Name$ Button1��5#      $*��$Зарегистрироваться��6&   $��24(�6C   $
��Bold"�6Q   ���	#FFE0E6E7���	#FFB4ABAB$��2q�6�   $��98,452,74,31q�6�   $	 0.558,0.885&�6�   $
��Auto��6�   ��6   Z� ��5$      �� 5%       �� 
TargetType" ��6&   .R�5&      ��  ��Name$ border��6!   #��z��61   #��z��6]   #��z��6�    J� SnapsToDevicePixels .  �� CornerRadius$ 20��6�   ._�5'      ��  ��Name$ contentPresenter��6/   #��z��6I   #��z��6}    �� ContentStringFormat �� ContentStringFormat# z 6�   |�.  #��z��6�   #��z��62  ��. # z 6p  #��z��6�  5(       �� Triggers 5)      P� 5*       P� Property �@��IsDefaulted P� Value . . �5+   "   ��  �� 
TargetName$ border��6@    �� Value# � ��6T    �� Property � ��5,      P� 5-       �@J�IsMouseOver . . �5.   "   �� $ border��6?    � �� �	#FF8738386S   �� 5/   "   $ border��6@    � �� �	#FF333F476T   50      P� 51       � �� . . �52   "   �� $ border��6?    � �� �	#FF7A48486S   �� 53   "   $ border��6@    � �� �	#FF333F476T   54      P� 55       �@d�	IsChecked$
 TrueY�6H   . �56   "   �� $ border��6?    � �� �	#FFBCDDEE6S   �� 57   "   $ border��6@    � �� �	#FF245A836T   58      P� 59       � {� .  . �5:   "   �� $ border��6?    � �� �	#FFF4F4F46S   �� 5;   "   $ border��6@    � �� �	#FFADB2B56T   �� 5<   "    � �� �	#FF8383836?   5=      5?      5D      � 5E      -   5         �Name$ ToolBar��5E      
������ $
��Left=�6=   $��27��6X   $	��Top=�6d   $	��458��6|   .+�5F      �� -   5         ��Name$ closer��5F      $
��Left=�6(   $��12��6C   $��
431,10,0,0q�6O   $	��Top=�6c   $��17��6{   $	 0.481,0.286&�6�   ��5G      �� 5H        �� 
TargetType"  ��6    .�5I      ��  � ��$ крест.png��67   ! �� Triggers! 5J      P� 5K   "    �@J�IsMouseOver . . �5L   &   ��  � ��$&  крестнаведение.png��6?   5N   #   5P      5R      �� 5S      -   5        $ remover��5S      $
��Left=�6)   $��9��6D   $��
403,13,0,0q�6O   $	��Top=�6c   $��16��6{   ��5T      �� 5U      "  ��6    .�5V      ��  � ��$ тире.png��67   ! 5W      P� 5X   "    �@J�IsMouseOver . . �5Y   &   ��  � ��$$ тиренаведение.png��6?   5[   #   5]      5_      5a      �� 5b      $��Регистрация��6   $��./#Futurespore Cyrillic)�6*   $
��Left=�6O   $��	98,92,0,0q�6j   $	��Top=�6}   ���	#FFE0E6E7$��83��6�   $	��304��6�   $��44(�6�   ��6�   Z� 6�   �� 5c      -   5        $	 TB3��5c      $
��Left=�6#   
 �����$ Отчество��6t   $��Times New Roman)�6�   $��40��6�   $��125,289,0,0q�6�   $
 Wrap=�6�   $	��Top=�6�   $	��230��6	  ���	#FFE0E6E7$��25(�6B  $ 15��6j  ��6  Z� 6y  �� 5d      -	   5        $	 TB4��5d      $
��Left=�6#   
 �����$ 
Логин��6t   $��Times New Roman)�6�   $��41��6�   $��125,338,0,0q�6�   $
 Wrap=�6�   $	��Top=�6�   $	��230��6  ���	#FFE0E6E7$��25(�6?  $ 15��6M  ��6  Z� 6M  �� 5e      -
   5        $	 TB5��5e      $
��Left=�6#   
 �����$ Пароль��6t   $��Times New Roman)�6�   $��42��6�   $��125,379,0,0q�6�   $
 Wrap=�6�   $	��Top=�6�   $	��230��6  ���	#FFE0E6E7$��25(�6@  $ 15��6N  ��6  Z� 6N  5g      5i      5j      