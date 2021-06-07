##################################################################
## PS Script to create a custom WPF content menu for a tray app ##
##################################################################

# Add assemblies
Add-Type -AssemblyName System.Windows.Forms,PresentationFramework

# Create a WinForms application context
$appContext = New-Object System.Windows.Forms.ApplicationContext

# Create an icon from base64
$Base64 = "Qk2KEAAAAAAAAIoAAAB8AAAAIAAAACAAAAABACAAAwAAAAAQAADDDgAAww4AAAAAAAAAAAAAAAD/AAD/AAD/AAAAAAAA/0JHUnOPwvUoUbgeFR6F6wEzMzMTZmZmJmZmZgaZmZkJPQrXAyhcjzIAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAA/wAAAP8AAAD/AAAA9wAAAOEAAACxAAAAbAAAABgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFkAAABWAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/AAAB/wADCP8AAgb/AAEC/wAAAP8AAAD/AAAA8wAAAI4AAAALAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABZAAAA+QAAAPcAAABVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8BBg//C1rh/wta4f8KUs3/CECf/wQeTP8AAQP/AAAA/wAAANAAAABEAAAABgAAAAAAAAAAAAAAWwAAAPgAAAD/AAAA/wAAAPcAAABWAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA+AEFDP8LXuz/DGP3/w1k9/8MY/f/DGP3/wpQx/8CEi3/AAAA/wAAAP8AAADmAAAAnAAAAHsAAAD6AQAA/3pfAP94XQD/AQEA/wAAAPkAAABaAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADhAAED/wpX2v8NZPf/N7z+/zOy/f8ikfv/Dmf3/wxd5v8eJTH/Dg0M/wEBAf8AAAD/AAAA/wIBAP97YAD/8rwA//G7AP93XAD/AQEA/wAAAPkAAABWAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALMAAAD/CEKl/wxj9/8zs/3/Pcj//z3I//81uP3/NmKh/1lUS/9YU0n/SkY9/y0qJf8ODQv/e2AB//G7AP/yvAD/8rwA//C6AP92XAD/AAAA/wAAAOoAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcQAAAP8EIVL/DGP3/yOT+/89yP//Pcf+/0uOo/9ZVUv/WVRK/1lUSv9ZVEr/WVRK/6mnof/32nj/8rwB//K8AP/yvAD/8rwA/+q2AP8HBgD/AAAA7gAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfAAAA+AADCP8KUs3/D2n3/za5/f9LjKH/WVVL/1lUSv9ZVEr/WVRK/1lUSv+rqKP//v7+///+/f/43Hv/8rwA//K8AP/yvAD/67YA/wgGAP8AAADuAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACTAAAA/wISLv8MXun/NmGf/1lUSv9ZVEr/WVRK/1lUSv9aVUv/qqij/////////////////////v/43Hv/8rwB//K8AP/rtgD/CAYA/wAAAO4AAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAAADaAAAA/yAoNv9ZVEv/WVRK/1lUSv9ZVEr/WlVL/6qnnv/+/v7////////////////////////+/f/423j/8rwB/+u2AP8IBgD/AAAA7gAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEsAAAD/ERAO/1lUSv9ZVEr/WVRK/1lUSv+qpZr/88Yr//zuvv/////////////////////////////+/P/423j/67cB/wgGAP8AAADuAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADQAAAOwBAQH/TUhA/1lUSv9ZVEr/q6ij//7+/v/77Lb/88EX//ztu//////////////////////////////+/f/x13v/CAYB/wAAAPsAAABcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAogAAAP8uKyb/WlVL/6qoo//////////////////77Ln/88EX//zuvv////////////////////////////////+CgoL/AQEB/wAAAPkAAABaAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB6AAAA/xIRD/+qp6L//v7+///////////////////////77Lb/88EX//zuv/////////////////////////////39/f9/f3//AQEB/wAAAPkAAABbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWwAAAPgBAQD/fGAC//fdgv////7////////////////////////////767X/88EX//zuvv////////////////////////////39/f9/f3//AQEB/wAAAPgAAABaAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF0AAAD6AQAA/3tgAP/xvAD/8rwC//negf////7////////////////////////////77Lb/9Mcr///+/P////////////////////////////39/f+CgoL/AQEB/wAAAPcAAABTAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABZAAAA+wEBAP97XwD/8rwA//K8AP/yvAD/8rwA//negf////7//////////////////////////////vv//////////////////f39//39/f////////////////9/f3//AQEB/wAAAPQAAAA2AAAAAAAAAAAAAAAAAAAAAAAAAF0AAAD8AQEA/39jAP/yvAD/8rwA//K8AP/yvAD/8rwC//neg/////7//////////////////////////////////Pz8/8rKyv+urq7/rq6u/8zMzP/9/f3///////z8/P9nZ2f/AAAA/wAAAOAAAAASAAAAAAAAAAAAAAAAAAAAAAAAAGAAAAD8AgEA/39iAP/yvAD/8rwA//K8AP/yvAD/8r0D//neg/////7////////////////////////////Ly8v/s7Oz/8jIyP/IyMj/srKy/8zMzP////////////X19f83Nzf/AAAA/wAAAKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF8AAAD6AgIA/35iAP/rtgD/67YA/+u2AP/rtgD/67cC//HYgf////7//////////////////Pz8/6ysrP/Jycn/zMzM/8zMzP/IyMj/rq6u//39/f///////////9fX1/8LCwv/AAAA+wAAADwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF4AAAD6AQEA/wgGAP8IBgD/CAYA/wgGAP8IBgD/CAYB/4SEhP/+/v7////////////8/Pz/q6ur/8nJyf/MzMz/zMzM/8jIyP+urq7//f39/////////////////39/f/8AAAD/AAAAwQAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF8AAAD1AAAA9gAAAPYAAAD2AAAA9gAAAPcAAAD9AgIC/4eHh//+/v7////////////Jycn/s7Oz/8nJyf/Jycn/s7Oz/8rKyv//////////////////////8PDx/xAPFf8AAAD8AAAANgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAARAAAAEQAAABEAAAARAAAAEQAAAGAAAAD7AwMD/4eHh//+/v7///////r6+v/Jycn/q6ur/6ysrP/Ly8v//Pz8//////////////////39//+Ng/P/Ewpx/wAAAP8AAACjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF0AAAD7AgIC/4SEhP/////////////////8/Pz//Pz8///////////////////////9/f//ioHz/yUT6P8fD8b/AAAC/wAAAOwAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF8AAAD7AQEB/4SEhP/+/v7//////////////////////////////////v7//46F8/8kEuj/JBLo/yQS5/8HAyv/AAAA/gAAAEEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAGAAAAD7AgIC/3Fxcf/6+vr///////////////////////7+//+RiPP/JRPo/yQS6P8kEuj/JBLo/xEIbv8AAAD/AAAAhAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAFkAAAD3AQEB/0FBQf/f39/////////////+/v//kIfz/yUT6P8kEuj/JBLo/yQS6P8kEuj/GQyg/wAAAP8AAAC2AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD0AAADjAAAA/w4ODv+Dg4P/8fHz/42E8/8lE+j/JBLo/yQS6P8kEuj/JBLo/yQS6P8eD8H/AAAB/wAAAN4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABUAAACpAAAA/QAAAP8QDxf/Ewp2/yAQzf8kEuj/JBLo/yQS6P8kEuj/JBLo/yEQ0/8BAAX/AAAA9gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAABCAAAAygAAAP4AAAD/AQEH/wgEMv8RCXD/Gg2n/yAQzf8iEd3/IRDT/wEBB/8AAAD+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAQQAAAKgAAAD0AAAA/wAAAP8AAAD/AAAD/wIBCv8CAQ7/AAAB/wAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAABJAAAAhgAAALcAAADfAAAA9gAAAP8AAAD/AAAA/w=="
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
$TrayIcon.Text = "Website Launcher"

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
        <MenuItem Name="Open" Header="Website Launcher" FontWeight="Bold" Style="{StaticResource MainMenuitem}"/>
        <Separator Width="240" Height="0.5" />

        <MenuItem Name="M365Portals" Header="M365 Portals" Style="{StaticResource SubMenuParentitem}"/> 
        <MenuItem Name="BunningsApps" Header="Bunnings Apps" Style="{StaticResource SubMenuParentitem}"/>
        <MenuItem Name="ITProTools" Header="IT Pro Tools" Style="{StaticResource SubMenuParentitem}"/>
        <MenuItem Name="Options" Header="Options" Style="{StaticResource SubMenuParentitem}"/>
        <MenuItem Name="Exit" Header="Exit" Style="{StaticResource MainMenuitem}"/>
        <Line Margin="3"/>
        <Popup Name="PortalPopup" PopupAnimation="Fade" Focusable="True" Placement="Right" PlacementTarget="{Binding ElementName=M365Portals}" Style="{StaticResource Popup}" HorizontalOffset="-5">
            <Border Background="#333333" BorderBrush="White" BorderThickness="0.0">
            <StackPanel >
                <Line Margin="3"/>
                <MenuItem Name="AdminPortal" Header="Admin Portal" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="AzurePortal" Header="Azure Portal" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="TeamsPortal" Header="Teams Admin Portal" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="MTRMSPortal" Header="MTR MS Portal" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="MEMCMConsole" Header="MEMCM Console" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="AppsAdmConsole" Header="Apps ADM Console" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="MSStoreforBus" Header="MS Store 4 Business" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="VLSCPortal" Header="VLSC Portal" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="VSSubscrip" Header="VS Subscription" Style="{StaticResource SubMenuitem}"/>
                <Line Margin="3"/>
            </StackPanel>
            </Border>
        </Popup>
        <Popup Name="BunPortalPopup" PopupAnimation="Fade" Focusable="True" Placement="Right" PlacementTarget="{Binding ElementName=BunningsApps}" Style="{StaticResource Popup}" HorizontalOffset="-5">
            <Border Background="#333333" BorderBrush="White" BorderThickness="0.0">
            <StackPanel >
                <Line Margin="3"/>
                <MenuItem Name="Intranet" Header="Intranet" Style="{StaticResource SubMenuitem}"/>
                <Line Margin="3"/>
            </StackPanel>
            </Border>
        </Popup>
        <Popup Name="ITProToolsPopup" PopupAnimation="Fade" Focusable="True" Placement="Right" PlacementTarget="{Binding ElementName=ITProTools}" Style="{StaticResource Popup}" HorizontalOffset="-5">
            <Border Background="#333333" BorderBrush="White" BorderThickness="0.0">
            <StackPanel >
                <Line Margin="3"/>
                <MenuItem Name="AdaptCards" Header="Adaptive Cards" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="GPOSearch" Header="Group Policy Search" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="MSAdoption" Header="MS Adoption Portal" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="PBIPlaygnd" Header="PowerBi Playground" Style="{StaticResource SubMenuitem}"/>
                <MenuItem Name="MSFasttrk" Header="MS FastTrack" Style="{StaticResource SubMenuitem}"/>
                <Line Margin="3"/>
            </StackPanel>
            </Border>
        </Popup>        
        <Popup Name="Popup" PopupAnimation="Fade" Focusable="True" Placement="Right" PlacementTarget="{Binding ElementName=Options}" Style="{StaticResource Popup}" HorizontalOffset="-5">
            <Border Background="#333333" BorderBrush="White" BorderThickness="0.9">
            <StackPanel >
                <Line Margin="3"/>
                <MenuItem Name="msportals" Header="MSPortals.io" Style="{StaticResource SubMenuitem}"/>
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
    $CM.AdminPortal.Add_Click({
        Start-Process MSEdge "https://admin.microsoft.com/"
        $CM.ContextMenu.Close()
    })
    # Open Azure Portal
    $CM.AzurePortal.Add_Click({
        Start-Process MSEdge "https://portal.azure.com/"
        $CM.ContextMenu.Close()
    })
    # Open M365 Teams Admin Portal
    $CM.TeamsPortal.Add_Click({
        Start-Process MSEdge "https://admin.teams.microsoft.com/dashboard"
        $CM.ContextMenu.Close()
    })
    # Open MTR Managed Services Portal MEMCMConsole
    $CM.MTRMSPortal.Add_Click({
        Start-Process MSEdge "https://portal.rooms.microsoft.com/"
        $CM.ContextMenu.Close()
    })
    # Open MEMCM Console
    $CM.MEMCMConsole.Add_Click({
        Start-Process MSEdge "https://endpoint.microsoft.com/"
        $CM.ContextMenu.Close()
    })
    # Open Apps Admin Console
    $CM.AppsAdmConsole.Add_Click({
        Start-Process MSEdge "https://config.office.com/"
        $CM.ContextMenu.Close()
    })
    # Open MS Store for Business Portal
    $CM.MSStoreforBus.Add_Click({
        Start-Process MSEdge "https://businessstore.microsoft.com/"
        $CM.ContextMenu.Close()
    })
    # Open VLSC Portal 
    $CM.VLSCPortal.Add_Click({
        Start-Process MSEdge "https://www.microsoft.com/Licensing/servicecenter/"
        $CM.ContextMenu.Close()
    })
    # Open Visual Studio Subscription
    $CM.VSSubscrip.Add_Click({
        Start-Process MSEdge "https://my.visualstudio.com/"
        $CM.ContextMenu.Close()
    })
    # Open Adaptive Cards Portal
    $CM.AdaptCards.Add_Click({
        Start-Process MSEdge "https://amdesigner.azurewebsites.net/"
        $CM.ContextMenu.Close()
    })    
    # Open Group Policy Search
    $CM.GPOSearch.Add_Click({
        Start-Process MSEdge "https://gpsearch.azurewebsites.net/"
        $CM.ContextMenu.Close()
    })    
    # Open MS Adoption Portal
    $CM.MSAdoption.Add_Click({
        Start-Process MSEdge "https://adoption.microsoft.com/"
        $CM.ContextMenu.Close()
    })    
    # Open Power BI Playground
    $CM.PBIPlaygnd.Add_Click({
        Start-Process MSEdge "https://microsoft.github.io/PowerBI-JavaScript/demo/v2-demo/index.html"
        $CM.ContextMenu.Close()
    })    
    # Open Microsoft FastTrack Portal
    $CM.MSFasttrk.Add_Click({
        Start-Process MSEdge "https://fasttrack.microsoft.com/"
        $CM.ContextMenu.Close()
    })    
    # Open MSPortals.io
    $CM.msportals.Add_Click({
        Start-Process MSEdge "https://msportals.io/"
        $CM.ContextMenu.Close()
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
