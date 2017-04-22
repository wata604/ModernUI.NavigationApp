Imports FirstFloor.ModernUI.Presentation
Imports System.ComponentModel

''' <summary>
''' A simple view model for configuring theme, font and accent colors.
''' </summary>
Public Class AppearanceViewModel
    Inherits NotifyPropertyChanged

    Private Const FontSmall As String = "small"
    Private Const FontLarge As String = "large"

    ' 9 accent colors from metro design principles
    Private _accentColors As Color() = New Color() {
        Color.FromRgb(&H33, &H99, &HFF),    ' blue
        Color.FromRgb(&H0, &HAB, &HA9),    ' teal
        Color.FromRgb(&H33, &H99, &H33),    ' green
        Color.FromRgb(&H8C, &HBF, &H26),    ' lime
        Color.FromRgb(&HF0, &H96, &H9),    ' orange
        Color.FromRgb(&HFF, &H45, &H0),    ' red
        Color.FromRgb(&HFF, &H0, &H97),    ' magenta
        Color.FromRgb(&HA2, &H0, &HFF)     ' purple
        }

    Private _selectedAccentColor As Color
    Private _themes As LinkCollection = New LinkCollection()
    Private _selectedTheme As Link
    Private _selectedFontSize As String

    Public Sub New()
        ' add the default themes
        Me._themes.Add(New Link With {.DisplayName = "dark", .Source = AppearanceManager.DarkThemeSource})
        Me._themes.Add(New Link With {.DisplayName = "light", .Source = AppearanceManager.LightThemeSource})

        Me.SelectedFontSize = IIf(AppearanceManager.Current.FontSize = FontSize.Large, FontLarge, FontSmall)
        SyncThemeAndColor()

        AddHandler AppearanceManager.Current.PropertyChanged, AddressOf OnAppearanceManagerPropertyChanged
    End Sub

    Private Sub SyncThemeAndColor()
        ' synchronizes the selected viewmodel theme with the actual theme used by the appearance manager.
        Me.SelectedTheme = Me._themes.FirstOrDefault(Function(l) l.Source.Equals(AppearanceManager.Current.ThemeSource))

        ' and make sure accent color is up-to-date
        Me.SelectedAccentColor = AppearanceManager.Current.AccentColor
    End Sub

    Private Sub OnAppearanceManagerPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        If e.PropertyName = "ThemeSource" OrElse e.PropertyName = "AccentColor" Then SyncThemeAndColor()
    End Sub

    Public ReadOnly Property Themes As LinkCollection
        Get
            Return Me._themes
        End Get
    End Property

    Public ReadOnly Property FontSizes As String()
        Get
            Return New String() {FontSmall, FontLarge}
        End Get
    End Property

    Public ReadOnly Property AccentColors As Color()
        Get
            Return Me._accentColors
        End Get
    End Property

    Public Property SelectedTheme As Link
        Get
            Return Me._selectedTheme
        End Get
        Set(value As Link)
            If Me._selectedTheme IsNot value Then
                Me._selectedTheme = value
                OnPropertyChanged("SelectedTheme")

                ' and update the actual theme
                AppearanceManager.Current.ThemeSource = value.Source
            End If
        End Set
    End Property

    Public Property SelectedFontSize As String
        Get
            Return Me._selectedFontSize
        End Get
        Set(value As String)
            If Me._selectedFontSize <> value Then
                Me._selectedFontSize = value
                OnPropertyChanged("SelectedFontSize")

                AppearanceManager.Current.FontSize = IIf(value = FontLarge, FontSize.Large, FontSize.Small)
            End If
        End Set
    End Property

    Public Property SelectedAccentColor As Color
        Get
            Return Me._selectedAccentColor
        End Get
        Set(value As Color)
            If Me._selectedAccentColor <> value Then
                Me._selectedAccentColor = value
                OnPropertyChanged("SelectedAccentColor")

                AppearanceManager.Current.AccentColor = value
            End If
        End Set
    End Property
End Class
