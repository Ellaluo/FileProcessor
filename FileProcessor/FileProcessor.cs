using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using Dapper;

namespace FileProcessor
{
    public class AttachmentModel
    {
        public byte[] ContentAsBytes { get; set; }
        public int RecordId { get; set; }
        public string FileType { get; set; }
    }

    public class FileProcessor
    {
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.Default;
                graphics.InterpolationMode = InterpolationMode.Bicubic;
                graphics.SmoothingMode = SmoothingMode.Default;
                graphics.PixelOffsetMode = PixelOffsetMode.Default;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public byte[] ImageToByteArray(Image image)
        {
            var result = new MemoryStream();

            image.Save(result, ImageFormat.Png);

            return result.ToArray();
        }

        public static Image ConvertBinaryFile(AttachmentModel attachment, int width, int height, bool raw)
        {
            if (!raw)
            {
                var img = ByteArrayToImage(attachment.ContentAsBytes);
                img = ResizeImage(img, width, height);
                return img;
            }
            else
            {
                var img = ByteArrayToImage(attachment.ContentAsBytes);
                height = img.Height;
                width = img.Width;
                img = ResizeImage(img, width, height);
                return img;
            }
        }

        public static Image ByteArrayToImage(byte[] bytes)
        {
            try
            {
                Image result;
                using (var ms = new MemoryStream(bytes))
                {
                    result = Image.FromStream(ms);
                }
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }


        public static void Main(string[] args)
        {
            string connectionString = "Server=;Database=;User Id=;Password=;MultipleActiveResultSets=true";
            int companyID = 567;
            var sqlBase = @"
                SELECT FileType, att.[Attachment] AS ContentAsBytes,
                att.[RecordID] AS RecordID 
                FROM dbo.[Attachment] att
				INNER JOIN dbo.[Risk] rsk ON rsk.RecordID = att.RiskID
                WHERE rsk.CompanyID IN 
                ((SELECT ChildCompanyID FROM MultiLink_CompanyLink WHERE ParentCompanyID = @companyID)
                UNION
                (SELECT RecordID FROM Company WHERE RecordID = @companyID))
                AND (FileType = 'jpg' or FileType = 'png' or FileType = 'pdf')";

            using (var connection = new SqlConnection(connectionString))
            {
                var attachmenList = connection.QueryAsync<AttachmentModel>(
                    sqlBase,
                    new { companyID }
                );

                foreach (var attachment in attachmenList.Result)
                {
                    int id = attachment.RecordId;
                    Byte[] byteArray = attachment.ContentAsBytes;
                    string fileType = attachment.FileType;
                    if (fileType.ToLower() == "png" | fileType.ToLower() == "jpg")
                    {
                        ImageFormat imageFormat;
                        switch (fileType.ToLower())
                        {
                            case "jpg":
                                imageFormat = ImageFormat.Jpeg;
                                GenerateImageFile(imageFormat, attachment, companyID, id);
                                break;
                            case "png":
                                imageFormat = ImageFormat.Png;
                                GenerateImageFile(imageFormat, attachment, companyID, id);
                                break;
                        }
                    }

                    else if (fileType.ToLower() == "pdf")
                    {
                        File.WriteAllBytes($@"{Path}/Attachment/{companyID}/Pdf/{id}.pdf", byteArray);
                    }
                }
            }
        }

        public const string Path = "C:/Code/FileProcessor";
        public static void GenerateImageFile(ImageFormat imageFormat, AttachmentModel attachment, int companyID, int id)
        {
            bool pathExists = Directory.Exists(Path);
            if (!pathExists)
            {
                throw new Exception("Please check your project path and custmise the path in the script.");
            }
            string attachmentPath = $@"{Path}/Attachment";
            bool attachmentPathExists = Directory.Exists(attachmentPath);

            if (!attachmentPathExists)
            {
                Directory.CreateDirectory(attachmentPath);
            }
            string smallPath = $@"{Path}/Attachment/{companyID}/Small";
            bool smallPathExists = Directory.Exists(smallPath);
            if (!smallPathExists)
            {
                Directory.CreateDirectory(smallPath);
            }
            string largePath = $@"{Path}/Attachment/{companyID}/Large";
            bool largePathExists = Directory.Exists(largePath);
            if (!largePathExists)
            {
                Directory.CreateDirectory(largePath);
            }
            string rawPath = $@"{Path}/Attachment/{companyID}/Raw";
            bool rawPathExists = Directory.Exists(rawPath);
            if (!rawPathExists)
            {
                Directory.CreateDirectory(rawPath);
            }
            string pdfPath = $@"{Path}/Attachment/{companyID}/Pdf";
            bool pdfPathExists = Directory.Exists(pdfPath);
            if (!pdfPathExists)
            {
                Directory.CreateDirectory(pdfPath);
            }

            int smallWidth = 200;
            int smallheight = 130;
            var smallImage = ConvertBinaryFile(attachment, smallWidth, smallheight, false);
            smallImage.Save($@"{Path}/Attachment/{companyID}/Small/{id}_200x130.{imageFormat}", imageFormat);

            int largeWidth = 1024;
            int largeheight = 666;
            var largeImage = ConvertBinaryFile(attachment, largeWidth, largeheight, false);
            largeImage.Save($@"{Path}/Attachment/{companyID}/Large/{id}_1024x666.{imageFormat}", imageFormat);

            int rawWidth = 0;
            int rawheight = 0;
            var rawImage = ConvertBinaryFile(attachment, rawWidth, rawheight, true);
            rawImage.Save($@"{Path}/Attachment/{companyID}/Raw/{id}_raw.{imageFormat}", imageFormat);
        }
    }
}
