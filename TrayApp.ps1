##################################################################
## PS Script to create a custom WPF content menu for a tray app ##
##################################################################

# Add assemblies
Add-Type -AssemblyName System.Windows.Forms,PresentationFramework

# Create a WinForms application context
$appContext = New-Object System.Windows.Forms.ApplicationContext

# Create an icon from base64
$Base64 = "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"
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($Base64)
$bitmap.EndInit()
$bitmap.Freeze()
$image = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($bitmap.StreamSource)
$icon = [System.Drawing.Icon]::FromHandle($image.GetHicon())

# Create a notify icon
$script:TrayIcon = New-Object System.Windows.Forms.NotifyIcon
$TrayIcon.Icon = $icon 
$TrayIcon.Text = "My Cool App"

# Make PowerShell Disappear - Thanks Chrissy
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

# Function to create the context menu
Function Create-ContextMenu {

# Define the window in XAML
[xml]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ContextMenu" SizeToContent="WidthAndHeight" WindowStyle="None" AllowsTransparency="True" Topmost="True" BorderBrush="White" BorderThickness="0.6" >
    <Window.Resources>
        <Style x:Key="MainMenuitem" TargetType="MenuItem">
            <Setter Property="Height" Value="35"/>
            <Setter Property="Width" Value="250"/>
            <Setter Property="Foreground" Value="WhiteSmoke"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type MenuItem}" >
                        <Border x:Name="Border" Padding="35,10,10,5" BorderThickness="0" Margin="2,0,2,0">
                            <ContentPresenter ContentSource="Header" x:Name="HeaderHost" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter Property="Background" TargetName="Border" Value="#4C4C4C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SubMenuitem" TargetType="MenuItem">
            <Setter Property="Height" Value="35"/>
            <Setter Property="Width" Value="200"/>
            <Setter Property="Foreground" Value="WhiteSmoke"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type MenuItem}" >
                        <Border x:Name="Border" Padding="35,10,10,5" BorderThickness="0" Margin="2,0,2,0">
                            <ContentPresenter ContentSource="Header" x:Name="HeaderHost" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter Property="Background" TargetName="Border" Value="#4C4C4C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>   
        <Style x:Key="SubMenuParentitem" TargetType="MenuItem">
            <Setter Property="Height" Value="35"/>
            <Setter Property="Width" Value="250"/>
            <Setter Property="Foreground" Value="WhiteSmoke"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type MenuItem}" >
                        <Border x:Name="Border" Padding="35,10,10,5" BorderThickness="0" Margin="2,0,2,0">
                            <Grid VerticalAlignment="Center">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <ContentPresenter ContentSource="Header" x:Name="HeaderHost" RecognizesAccessKey="True"/>
                                <Path Name="SelectedArrow" Data="M 0 12 L 6 6 L 0 0" Stroke="White" Margin="190,0,0,0"  />
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter Property="Background" TargetName="Border" Value="#4C4C4C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="Popup" TargetType="{x:Type Popup}">
            <Setter Property="IsOpen" Value="True" />
            <Style.Triggers>
                <MultiDataTrigger>
                    <MultiDataTrigger.Conditions>
                        <Condition Binding="{Binding RelativeSource={RelativeSource Self},Path=PlacementTarget.IsMouseOver}" Value="False" />
                        <Condition Binding="{Binding RelativeSource={RelativeSource Self},Path=IsMouseOver}" Value="False" />
                    </MultiDataTrigger.Conditions>
                    <Setter Property="IsOpen" Value="False" />
                </MultiDataTrigger>
            </Style.Triggers>
        </Style>           
    </Window.Resources>
    
    <StackPanel Background="#333333" >
        <Line Margin="3"/>
        <MenuItem Name="Open" Header="Open My Cool App" FontWeight="Bold" Style="{StaticResource MainMenuitem}"/>
        <Separator Width="240" Height="0.5" />

        <MenuItem Name="M365Portals" Header="M365 Portals" Style="{StaticResource SubMenuParentitem}"/>
        <MenuItem Name="Options" Header="Options" Style="{StaticResource SubMenuParentitem}"/>
        <MenuItem Name="Exit" Header="Exit" Style="{StaticResource MainMenuitem}"/>
        <Line Margin="3"/>
        <Popup Name="PortalPopup" PopupAnimation="Fade" Focusable="True" Placement="Right" PlacementTarget="{Binding ElementName=M365Portals}" Style="{StaticResource Popup}" HorizontalOffset="-5">
            <Border Background="#333333" BorderBrush="White" BorderThickness="0.0">
            <StackPanel >
                <Line Margin="3"/>
                <MenuItem Name="AdminPortal" Header="Admin Portal" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="TeamsPortal" Header="Teams Admin Portal" Style="{StaticResource SubMenuitem}"/>
                <Line Margin="3"/>
            </StackPanel>
            </Border>
        </Popup>
        <Popup Name="Popup" PopupAnimation="Fade" Focusable="True" Placement="Left" PlacementTarget="{Binding ElementName=Options}" Style="{StaticResource Popup}" HorizontalOffset="-5">
            <Border Background="#333333" BorderBrush="White" BorderThickness="0.9">
            <StackPanel >
                <Line Margin="3"/>
                <MenuItem Name="ChangeColour" Header="Change Colour" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="ChangeSize" Header="Change Size" Style="{StaticResource SubMenuitem}"/>
                <Line Margin="3"/>
            </StackPanel>
            </Border>
        </Popup>
    </StackPanel>
</Window>
"@          
    
    # Create the Windows and elements
    $global:CM = @{}
    $CM.ContextMenu = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | 
        ForEach-Object -Process {
            $CM.$($_.Name) = $CM.ContextMenu.FindName($_.Name)
        }
    
    # Move the menu window to the correct location by the mouse cursor
    $CM.ContextMenu.Add_Loaded({
        [System.Drawing.Point]$point = [System.Windows.Forms.Control]::MousePosition
        $MousePosition = [System.Windows.Point]::new($point.X,$point.Y)
        $Transform = [System.Windows.PresentationSource]::FromVisual($this).CompositionTarget.TransformFromDevice
        $Mouse = $Transform.transform($MousePosition)
        $CM.ContextMenu.Top = ($Mouse.Y - $This.ActualHeight)
        $CM.ContextMenu.Left = ($Mouse.X - $This.ActualWidth)
    })

    # Don't leave the menu window open if the mouse leaves it
   $CM.ContextMenu.Add_MouseLeave({
       If ($CM.Popup.IsMouseOver -ne $true)
        {
            #$this.Close()
        }
    })

    # Open the app
    $CM.Open.Add_Click({
        Start-Process Notepad
    })

    # Open M365 Admin Portal
    $CM.TeamsPortal.Add_Click({
        Start-Process MSEdge "https://admin.teams.microsoft.com/dashboard"
    })
    # Change colour option
    $CM.ChangeColour.Add_Click({
        $Colour = "Green","Blue","Red","Orange","Brown" | Get-Random
        $This.Foreground = $Colour
    })

    # Change size option
    $CM.ChangeSize.Add_Click({
        $Size = "6","8","10","12","14" | Get-Random
        $This.FontSize = $Size
    })

    # Clean up on exit
    $CM.Exit.Add_Click({
        $CM.ContextMenu.Close()
        $TrayIcon.Dispose()
        $appContext.ExitThread()
        $appContext.Dispose()
    })

    # Display the menu
    $CM.ContextMenu.ShowDialog()

}

# Open content menu on icon right-click
$TrayIcon.Add_MouseDown({
    If ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
        If ($CM.ContextMenu)
        {
            $CM.ContextMenu.Close()
        }
        Create-ContextMenu
    }
})


$TrayIcon.Visible = $true

[void][System.Windows.Forms.Application]::Run($appContext)
