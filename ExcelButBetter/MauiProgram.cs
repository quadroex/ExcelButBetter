using Microsoft.Extensions.Logging;
using CommunityToolkit.Maui;
using Microsoft.Maui.LifecycleEvents;

namespace ExcelButBetter
{
    public static class MauiProgram
    {
        public static MauiApp CreateMauiApp()
        {
            var builder = MauiApp.CreateBuilder();
            builder
               .UseMauiApp<App>()
               .UseMauiCommunityToolkit()
               .ConfigureFonts(fonts =>
               {
                   fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                   fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
               })

               // hook for windows
               .ConfigureLifecycleEvents(events =>
               {
#if WINDOWS
                    events.AddWindows(windows => windows.OnWindowCreated(window =>
                    {
                        var handle = WinRT.Interop.WindowNative.GetWindowHandle(window);
                        var id = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(handle);
                        var appWindow = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(id);
                        
                        // danger
                        appWindow.Closing += async (sender, args) =>
                        {
                            // danger
                            args.Cancel = true;

                            var mainPage = ExcelButBetter.MainPage.Instance;
                            if (mainPage!= null)
                            {
                                bool canClose = await mainPage.AskToSaveIfDirty();
                                
                                if (canClose)
                                {
                                    // cycle danger
                                    Application.Current.Quit();
                                }
                            }
                        };
                    }));
#endif
               });

#if DEBUG
            builder.Logging.AddDebug();
#endif
            return builder.Build();
        }
    }
}
