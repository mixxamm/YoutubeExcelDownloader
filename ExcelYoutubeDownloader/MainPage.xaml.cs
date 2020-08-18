using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text.RegularExpressions;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.AccessCache;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using YoutubeExplode;
using YoutubeExplode.Videos.Streams;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace ExcelYoutubeDownloader
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        private void KiesBestandButton_Click(object sender, RoutedEventArgs e)
        {
            KiesBestandAsync();
        }

        private async void KiesBestandAsync()
        {
            var picker = new Windows.Storage.Pickers.FileOpenPicker();
            picker.ViewMode = Windows.Storage.Pickers.PickerViewMode.Thumbnail;
            picker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.Downloads;
            picker.FileTypeFilter.Add(".xls");
            picker.FileTypeFilter.Add(".xlsx");
            Windows.Storage.StorageFile file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                KiesBestandButton.Visibility = Visibility.Collapsed;
                // Application now has read/write access to the picked file
                // this.textBlock.Text = "Picked photo: " + file.Name;
                var stream = await file.OpenStreamForReadAsync();
                List<string> links = new List<string>();
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    TextBlockItems.Text = "Zoeken naar links...";
                    do
                    {
                        while (reader.Read())
                        {
                            for(int i = 0; i < reader.FieldCount; i++)
                            {
                                string link = reader.GetString(i);
                                if (link != null && link.Contains("https://") && (link.Contains("youtube.com") || link.Contains("youtu.be")))
                                {
                                    links.Add(link);
                                }
                            }
                        }
                    } while (reader.NextResult());
                    TextBlockItems.Text = $"{links.Count} items gevonden.";
                    var folderPicker = new Windows.Storage.Pickers.FolderPicker();
                    folderPicker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.Desktop;
                    folderPicker.FileTypeFilter.Add("*");

                    Windows.Storage.StorageFolder folder = await folderPicker.PickSingleFolderAsync();
                    if (folder != null)
                    {
                        // Application now has read/write access to all contents in the picked folder
                        // (including other sub-folder contents)
                        Windows.Storage.AccessCache.StorageApplicationPermissions.
                        FutureAccessList.AddOrReplace("PickedFolderToken", folder);
                        var youtube = new YoutubeClient();
                        TextBlockItems.Text = "Bezig met downloaden.";
                        int aantalLinks = links.Count;
                        int teller = 0;
                        foreach (string link in links)
                        {
                            try
                            {
                                var video = await youtube.Videos.GetAsync(link);
                                var manifest = await youtube.Videos.Streams.GetManifestAsync(video.Id);
                                var streamInfo = manifest.GetAudioOnly().WithHighestBitrate();
                                if (streamInfo != null)
                                {
                                    var soundStream = await youtube.Videos.Streams.GetAsync(streamInfo);
                                    var storageFolder = await StorageApplicationPermissions.FutureAccessList.GetFolderAsync("PickedFolderToken");
                                    var storageFile = await storageFolder.CreateFileAsync($"{Regex.Replace(video.Title, "[^a-zA-Z0-9_.]+", " ", RegexOptions.Compiled)}.{streamInfo.Container}", CreationCollisionOption.ReplaceExisting);

                                    long length = soundStream.AsInputStream().AsStreamForRead().Length;

                                    var writeStream = await storageFile.OpenStreamForWriteAsync();
                                    await soundStream.CopyToAsync(writeStream);
                                    teller++;
                                    DownloadProgress.Value = (teller + 0.0) / (aantalLinks + 0.0) * 100;
                                    TextBlockItems.Text = $"{teller}/{aantalLinks} liedjes gedownload.";
                                }
                            }
                            catch (Exception e)
                            {
                                Debug.WriteLine(e);
                            }
                        }
                    }
                    else
                    {
                        TextBlockItems.Text = "Downloaden geannuleerd.";
                    }

                }
            }
            else
            {
                // this.textBlock.Text = "Operation cancelled.";
            }
        }

        private static byte[] ReadFully(Stream input)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}
