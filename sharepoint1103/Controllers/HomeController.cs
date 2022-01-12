using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Web;
using System.Web.Mvc;

namespace sharepoint1103.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page. "+ DateTime.Now;
            UploadFile();
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "upload c:Test.pdf 3254454" + DateTime.Now;
            UploadFile();
            UploadFile2("Test.pdf", 3254454);
            return View();
        }

        public void CreateFolder()
        {
            string userName = "srv01@pttlab.onmicrosoft.com";
            string Password = "Abcd@1234";
            var securePassword = new SecureString();
            foreach (char c in Password)
            {
                securePassword.AppendChar(c);
            }
            using (var ctx = new ClientContext("https://pttlab.sharepoint.com/sites/NFGDemo"))
            {
                ctx.Credentials = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(userName, securePassword);
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                List byTitle = ctx.Web.Lists.GetByTitle("Documents");

                // New object of "ListItemCreationInformation" class
                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();

                // Below are options.
                // (1) File - This will create a file in the list or document library
                // (2) Folder - This will create a foder in list(if folder creation is enabled) or documnt library
                listItemCreationInformation.UnderlyingObjectType = FileSystemObjectType.Folder;

                // This will et the internal name/path of the file/folder
                listItemCreationInformation.LeafName = "NewFolderFromCSOM";

                ListItem listItem = byTitle.AddItem(listItemCreationInformation);

                // Set folder Name
                listItem["Title"] = "NewFolderFromCSOM";

                listItem.Update();
                ctx.ExecuteQuery();
            }
        }
        public void UploadFile()
        {
            string userName = "srv01@pttlab.onmicrosoft.com";
            string Password = "Abcd@1234";
            var securePassword = new SecureString();
            foreach (char c in Password)
            {
                securePassword.AppendChar(c);
            }
            using (var ctx = new ClientContext("https://pttlab.sharepoint.com/sites/NFGDemo"))
            {
                ctx.Credentials = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(userName, securePassword);
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                FileCreationInformation newFile = new FileCreationInformation();
                newFile.Content = System.IO.File.ReadAllBytes("C:\\test.txt");
                newFile.Url = @"test.txt";

                newFile.Overwrite = true;
                List byTitle = ctx.Web.Lists.GetByTitle("Documents");
                Folder folder = byTitle.RootFolder.Folders.GetByUrl("NewFolderFromCSOM");
                ctx.Load(folder);
                ctx.ExecuteQuery();
                Microsoft.SharePoint.Client.File uploadFile = folder.Files.Add(newFile);
                var user = uploadFile.CheckedOutByUser;

                ctx.Load(byTitle);
                ctx.Load(uploadFile);
                ctx.ExecuteQuery();

                //ctx.Load(uploadFile, f => f.ListItemAllFields);
                //ctx.ExecuteQuery();


                if (uploadFile.CheckOutType != CheckOutType.None)
                {
                    uploadFile.CheckIn("1st checkin", CheckinType.MajorCheckIn);
                }
                ctx.ExecuteQuery();
                Console.WriteLine("done");
            }
        }

        public void UploadFile(string filename)
        {
            string userName = "srv01@pttlab.onmicrosoft.com";
            string Password = "Abcd@1234";
            var securePassword = new SecureString();
            foreach (char c in Password)
            {
                securePassword.AppendChar(c);
            }
            using (var ctx = new ClientContext("https://pttlab.sharepoint.com/sites/NFGDemo"))
            {
                ctx.Credentials = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(userName, securePassword);
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                FileCreationInformation newFile = new FileCreationInformation();
                newFile.Content = System.IO.File.ReadAllBytes("C:\\test.txt");
                newFile.Url = @filename;

                newFile.Overwrite = true;
                List byTitle = ctx.Web.Lists.GetByTitle("Documents");
                Folder folder = byTitle.RootFolder.Folders.GetByUrl("NewFolderFromCSOM");
                ctx.Load(folder);
                ctx.ExecuteQuery();
                Microsoft.SharePoint.Client.File uploadFile = folder.Files.Add(newFile);
                var user = uploadFile.CheckedOutByUser;

                ctx.Load(byTitle);
                ctx.Load(uploadFile);
                ctx.ExecuteQuery();

                //ctx.Load(uploadFile, f => f.ListItemAllFields);
                //ctx.ExecuteQuery();


                if (uploadFile.CheckOutType != CheckOutType.None)
                {
                    uploadFile.CheckIn("1st checkin", CheckinType.MajorCheckIn);
                }
                ctx.ExecuteQuery();
                Console.WriteLine("done");
            }
        }
        public void UploadFile(string filename, int filesize = 5000000)
        {
            string userName = "srv01@pttlab.onmicrosoft.com";
            string Password = "Abcd@1234";
            var securePassword = new SecureString();
            foreach (char c in Password)
            {
                securePassword.AppendChar(c);
            }
            //FileCreationInformation newFile = new FileCreationInformation();
            //newFile.Content = System.IO.File.ReadAllBytes("D:\\testBigSize.txt");
            var ctx = new ClientContext("https://pttlab.sharepoint.com/sites/NFGDemo");
            ctx.Credentials = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(userName, securePassword);
            Web web = ctx.Web;
            ctx.Load(web);
            ctx.ExecuteQuery();
            string strFolder = "Documents/NewFolderFromCSOM";
            string SiteURL = "https://pttlab.sharepoint.com/sites/NFGDemo";
            string uniqueFileName = @"test.pdf";// ReplaceName(Path.GetFileName(attach.FileName));
            string siteFolderURL = SiteURL + "/" + strFolder + "/";
            string siteFileURL = siteFolderURL + uniqueFileName;


            UploadFileSlicePerSliceToFolder(ctx, strFolder, filename);
        }

        public void UploadFile2(string filename, int filesize)
        {

            string userName = "srv01@pttlab.onmicrosoft.com";
            //userName = "ryan.tai@pttlab.onmicrosoft.com";
            string Password = "Abcd@1234";
            var securePassword = new SecureString();
            foreach (char c in Password)
            {
                securePassword.AppendChar(c);
            }



            FileCreationInformation newFile = new FileCreationInformation();
            newFile.Content = System.IO.File.ReadAllBytes("c:\\test.pdf");

            string strFolder = "Documents/NewFolderFromCSOM";
            string SiteURL = "https://pttlab.sharepoint.com/sites/NFGDemo";
            string uniqueFileName = @filename;// ReplaceName(Path.GetFileName(attach.FileName));
            string siteFolderURL = SiteURL + "/" + strFolder + "/";
            string siteFileURL = siteFolderURL + uniqueFileName;

            using (var ctx = new ClientContext("https://pttlab.sharepoint.com/sites/NFGDemo"))
            {

                Guid uploadId = Guid.NewGuid();
                List byTitle = ctx.Web.Lists.GetByTitle("Documents");
                Folder folder = byTitle.RootFolder.Folders.GetByUrl("NewFolderFromCSOM");

                ctx.Credentials = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(userName, securePassword);
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();


                //Folder folder = ctx.Web.GetFolderByServerRelativeUrl(siteFolderURL);
                ctx.Load(folder);

                ctx.ExecuteQuery();
                Microsoft.SharePoint.Client.File uploadFile = null;
                int fileChunkSizeInMB = 2;
                int blockSize = fileChunkSizeInMB * 1024 * 1024;

                ClientResult<long> bytesUploaded = null;

                try
                {
                    FileStream fs = new FileStream("C:\\test.pdf", FileMode.Open, FileAccess.ReadWrite);
                    //byte[] file = System.IO.File.ReadAllBytes("D:\\testBigSize.txt");

                    using (BinaryReader br = new BinaryReader(fs))  //(Attach.InputStream))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        //List byTitle = ctx.Web.Lists.GetByTitle("Documents");
                        //Folder folder = byTitle.RootFolder.Folders.GetByUrl("NewFolderFromCSOM");

                        // Read data from file system in blocks. 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // You've reached the end of the file.
                            if (totalBytesRead == filesize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = @filename; //siteFileURL;
                                    fileInfo.Overwrite = true;
                                    //uploadFile = docs.RootFolder.Files.Add(fileInfo);
                                    uploadFile = folder.Files.Add(fileInfo);

                                    //uploadFile = ctx.Web.GetFileByServerRelativeUrl(siteFolderURL + uniqueFileName);
                                    // Start upload by uploading the first slice. 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice.
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        //ctx.Load(uploadFile);
                                        ctx.ExecuteQuery();
                                        // fileoffset is the pointer where the next slice will be added.
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // You can only start the upload once.
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to your file.
                                //uploadFile = ctx.Web.GetFileByServerRelativeUrl(docs.RootFolder.ServerRelativeUrl + System.IO.Path.AltDirectorySeparatorChar + uniqueFileName);
                                //uploadFile = ctx.Web.GetFileByServerRelativeUrl(siteFolderURL + uniqueFileName);
                                //uploadFile = ctx.Web.GetFileByServerRelativeUrl("/Documents/NewFolderFromCSOM/" + uniqueFileName);
                                var spurl = web.ServerRelativeUrl;
                                var url = string.Format("{0}/Shared%20Documents/NewFolderFromCSOM/{1}", spurl, uniqueFileName);
                                uploadFile = ctx.Web.GetFileByServerRelativeUrl(url);
                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload.
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.Load(uploadFile);
                                        ctx.ExecuteQuery();

                                        ctx.Load(uploadFile, f => f.ListItemAllFields);
                                        ctx.ExecuteQuery();
                                        // Return the file object for the uploaded file.
                                        //return uploadFile;
                                    }
                                }
                                else
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload.
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // Update fileoffset for the next slice.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                    }
                }
                catch (Exception ex)
                {
                    ViewBag.Message = "Your application description page. throw Exception " + DateTime.Now;
                    //WriteToLog.writetolog(string.Format("Runtime Exception {0} | {1} | {2}", ex.Message, ex.InnerException, ex.StackTrace));
                    //WriteToLog.writetolog(string.Format("SharePointHandler - public static Microsoft.SharePoint.Client.File UploadFileSlicePerSlice(string filePath, int appID, string folderName, int fileChunkSizeInMB = 3) \n Runtime Exception {0} | {1} | {2}", ex.Message, ex.InnerException, ex.StackTrace));
                }
            }
        }


        public Microsoft.SharePoint.Client.File UploadFileSlicePerSliceToFolder(ClientContext ctx, string serverRelativeFolderUrl, string fileName, int fileChunkSizeInMB = 3)
        {
            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();

            // Get the name of the file.
            string uniqueFileName = Path.GetFileName(fileName);

            // Get the folder to upload into.
            Folder uploadFolder = ctx.Web.GetFolderByServerRelativeUrl(serverRelativeFolderUrl);

            // Get the information about the folder that will hold the file.
            ctx.Load(uploadFolder);
            ctx.ExecuteQuery();

            // File object.
            Microsoft.SharePoint.Client.File uploadFile = null;

            // Calculate block size in bytes.
            int blockSize = fileChunkSizeInMB * 1024 * 1024;

            // Get the information about the folder that will hold the file.
            ctx.Load(uploadFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();


            // Get the size of the file.
            long fileSize = new FileInfo(fileName).Length;

            if (fileSize <= blockSize)
            {
                // Use regular approach.
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = uniqueFileName;
                    fileInfo.Overwrite = true;
                    uploadFile = uploadFolder.Files.Add(fileInfo);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();
                    // Return the file object for the uploaded file.
                    return uploadFile;
                }
            }
            else
            {
                // Use large file upload approach.
                ClientResult<long> bytesUploaded = null;

                FileStream fs = null;
                try
                {
                    fs = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using (BinaryReader br = new BinaryReader(fs))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // Read data from file system in blocks.
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // You've reached the end of the file.
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = uniqueFileName;
                                    fileInfo.Overwrite = true;
                                    uploadFile = uploadFolder.Files.Add(fileInfo);

                                    // Start upload by uploading the first slice.
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice.
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        ctx.ExecuteQuery();
                                        // fileoffset is the pointer where the next slice will be added.
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // You can only start the upload once.
                                    first = false;
                                }
                            }
                            else
                            {
                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload.
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();

                                        // Return the file object for the uploaded file.
                                        return uploadFile;
                                    }
                                }
                                else
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload.
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // Update fileoffset for the next slice.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }
                        } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                    }
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Dispose();
                    }
                }
            }

            return null;
        }

    }
}