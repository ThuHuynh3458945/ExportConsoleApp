using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using SixLabors.ImageSharp.Formats.Png;
using System.IO;

namespace ExportConsoleApp
{
    public class ExportShipLabelUrlService
    {
        private static async Task<byte[]?> DownloadFileUrlAsync(string url)
        {
            byte[]? bytesFile;
            if (string.IsNullOrEmpty(url))
            {
                return Array.Empty<byte>();
            }
            try
            {
                using var client = new HttpClient();
                var response = await client.GetAsync(url);
                if (!response.IsSuccessStatusCode)
                {
                    return Array.Empty<byte>();
                }

                bytesFile = await response.Content.ReadAsByteArrayAsync();
            }
            catch
            {
                bytesFile = null;
            }

            return bytesFile;
        }

        private static byte[] GenerateShiplabel4X6(byte[] labelBytes)
        {
            // 1. Convert image from blob url
            // 2. Rotate image (image size: 6x4 -> 4x6)

            #region Define Document
            var stream = new MemoryStream();
            var writer = new PdfWriter(stream);
            var pdfDocument = new PdfDocument(writer);
            var pageSize = new PageSize(288f, 432f);
            pdfDocument.SetDefaultPageSize(pageSize);
            var document = new Document(pdfDocument);
            document.SetMargins(0f, 0f, 0f, 0f);
            #endregion

            #region Generate Shiplabel           
            ImageData imageData = ImageDataFactory.Create(labelBytes);
            var imagePdf = new iText.Layout.Element.Image(imageData);
            imagePdf.SetMargins(0f, 0f, 0f, 0f)
                .SetPadding(0f)
                .SetBorder(Border.NO_BORDER)
                ;
            document.Add(imagePdf);
            #endregion

            document.Close();
            return stream.ToArray();
        }

        public async Task<byte[]> ExportShipLabelUrl(string shipLabelUrl)
        {
            shipLabelUrl = "https://easypost-files.s3.us-west-2.amazonaws.com/files/postage_label/20230705/e61e8051f53f4448099b3cf0f70e16ec2b.png";
            var imgLabelBytes = await DownloadFileUrlAsync(shipLabelUrl);
            if (imgLabelBytes == null) return Array.Empty<byte>();

            //4x7 or 4x8
            using (var imgStream = new MemoryStream(imgLabelBytes))
            {
                var img = await Image.LoadAsync(imgStream);

                //get ratio page 
                var heightCheck = img.Height * 4 / 6;
                if (heightCheck > img.Width)
                {
                    var newHeight = img.Width * 6 / 4;
                    var width = img.Width;
                    var rect = new SixLabors.ImageSharp.Rectangle(0, 0, width, newHeight);
                    img.Mutate(i => i.Resize(width, newHeight).Crop(rect));

                    using (var ms = new MemoryStream())
                    {
                        await img.SaveAsync(ms, new PngEncoder());
                        imgLabelBytes = ms.ToArray();
                    }
                }  
            }

            //6x4
            //using (var ms = new MemoryStream(imgLabelBytes))
            //{
            //    //rotate the image -90 since it's provide in portrait mode
            //    var img = await Image.LoadAsync(ms);
            //    img.Mutate(x => x.Rotate(90));

            //    using (var stream = new MemoryStream(0))
            //    {
            //        await img.SaveAsync(stream, new PngEncoder());
            //        imgLabelBytes = stream.ToArray();
            //    }
            //}

            var result = GenerateShiplabel4X6(imgLabelBytes);
            return result;
        }
    }
}
