using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
//SourceCode.Hosting.Client is used to construct more advanced connection strings 
//this assembly can be added from the .NET references tab as SourceCode.HostClientAPI or from the file system at
//[Program Files]\[K2 Directory]\Bin\SourceCode.HostClientAPI.dll
using SourceCode.Hosting.Client;
//SourceCode.Workflow.CLient is used to interact with workflows and worklists programatically at runtime
//this assembly can be added from the .NET references tab as SourceCode.Workflow.Client or from the file system at
//[Program Files]\[K2 Directory]Bin\SourceCode.Workflow.Client.dll
using SourceCode.Workflow.Client;


namespace K2Documentation.Samples.WorkflowRuntime.WorkflowClientAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            //no implementation. Refer to the methods below for examples of K2 client API code
            //NOTE: this code will not run, it is intended only to show samples of specific API method calls.
        }

        //K2 connection opening samples
        public static void K2ConnectionSamples()
        {
            //you must instantiate the connection object
            Connection K2Conn = new Connection();

            //simple open using only a server name. (we are using localhost for this sample, in your example use your K2 Server name or NLB server-name)
            //This is the simplest way to open a connection for the current user (Active Directory credentials) on the default workflow port (5252)
            K2Conn.Open("[localhost]");

            //creating a more advanced connection with the connection string builder
            SourceCode.Hosting.Client.BaseAPI.SCConnectionStringBuilder K2ConnStringBuilder = new SourceCode.Hosting.Client.BaseAPI.SCConnectionStringBuilder();
            K2ConnStringBuilder.Host = "[localhost]"; //K2 server name or the name of the DNS entry pointing to the K2 Farm
            K2ConnStringBuilder.Authenticate = true; //specify whether to authenticate the user's credentials against the security provider. This is usually set to true
            K2ConnStringBuilder.UserID = "[username]"; //a specific username
            K2ConnStringBuilder.Password = "[password]";   //the user's password, unencrypted
            K2ConnStringBuilder.Port = 5252;  //if K2 was configured to run on a non-standard port, you can specify the port number here
            K2ConnStringBuilder.IsPrimaryLogin = true; //this is normally set to true, unless you are using cached security credentials
            K2ConnStringBuilder.SecurityLabelName = "K2"; //if using a different security provider, specify the label of the security provider here
            //opening a K2 connection with the advanced connection string
            K2Conn.Open(K2ConnStringBuilder.ToString());

            //closing a connection: you must always close the K2 connection when you are done with it
            K2Conn.Close();

            //to ensure that the connection is closed and disposed, you can also wrap the connection into a using statement, like this:
            using (Connection K2Conn2 = new Connection())
            {
                K2Conn2.Open("[servername]");
                //do something with the connection
                //the connection will be closed and disposed when the using statement is done
            }
        }


        //starting a workflow sample
        public static void StartWorkflowSamples()
        {
            //you must open a K2 connection first. We will wrap the K2 connection into a using statement to ensure that it is closed
            using (Connection K2Conn = new Connection())
            {
                K2Conn.Open("[servername]");
                //first, create a new process instance. This only loads the process into memory, it does not start it yet.
                //Note: if the user does not have rights to start the workflow, an error will be thrown here
                ProcessInstance K2Proc = K2Conn.CreateProcessInstance(@"[ProjectName]\[Workflow Name]"); //specify the full name of the workflow
                //now that we have the process as an object in memory, we can set some properties before starting the process
                K2Proc.Folio = "[ProcessFolio]";
                K2Proc.Priority = 1;
                //datafields should be accessed by name. Be aware of data types when setting values, and ensure that the data field names are spelled correctly!
                K2Proc.DataFields["[String DataField Name]"].Value = "[somevalue]";
                K2Proc.DataFields["[Integer Datafield Name]"].Value = 1;
                K2Proc.DataFields["[Date Datafield Name]"].Value = DateTime.Now;
                //XML fields are set using an XML-formatted string
                XmlDocument xmlDoc= new XmlDocument(); //TODO: set up the XML document as required, or construct a valid XML string somehow
                K2Proc.XmlFields["[XML DataField Name]"].Value = xmlDoc.ToString();
                //once you have set all the necessary values, start the process
                K2Conn.StartProcessInstance(K2Proc);

                //you can also start a process synchronously with the Sync override. This approach is used when you need to wait for the 
                //first client event/wait-state in the workflow before returning the StartProcessInstance call, most commonly with 
                //page-flow solutions or with automated testing code. This does take longer to return, so don't use it unless you have a specific reason to wait for the first wait-state or client event
                K2Conn.StartProcessInstance(K2Proc,true);   
            }
        }

        //opening worklist sample
        public static void OpenWorklistSamples()
        {
            //you must open a K2 connection first. We will wrap the K2 connection into a using statement to ensure that it is closed
            using (Connection K2Conn = new Connection())
            {
                K2Conn.Open("[servername]");
                //retrieve the entire worklist for the current user
                Worklist K2WList = K2Conn.OpenWorklist();

                //or retrieve the worklist for a specific platform ("platform" not really used that often anymore). The default platform for worklist items is ASP
                Worklist K2WList1 = K2Conn.OpenWorklist("ASP");

                //or retrieve the worklist for a specific platform and a user managed by the current user. 
                //This requires that the reporting structure is set up correctly in the underlying user provider (Active Directory or custom user manager)
                Worklist K2WList2 = K2Conn.OpenWorklist("ASP",@"[domain\username]");

                //you can also retrieve the worklist with a criteria filter. OpenWorklist can be an expensive operation, so use criteria to perform server-side filtering of the worklist 
                //for large worklists you should always use criteria to reduce the size of the worklist that is returned to the client
                WorklistCriteria K2Crit = new WorklistCriteria();
                //you can set up various combinations of filters and sorts in the criteria object. Refer to the product documentation for more samples
                //here, we are filtering for all workflows for the project K2Learning where the priority equals 1. We want to sort the task list by workflow start date
                K2Crit.AddFilterField(WCField.ProcessFolder, WCCompare.Equal, "[FolderName]");
                K2Crit.AddFilterField(WCLogical.And, WCField.ProcessPriority, WCCompare.Equal, 1);
                K2Crit.AddSortField(WCField.ProcessStartDate, WCSortOrder.Descending);
                Worklist K2WList3 = K2Conn.OpenWorklist(K2Crit);

                int numberOfTasks = K2WList.TotalCount;

                //once you have retrieved the worklist, you can iterate over the worklist items in the worklist
                foreach (WorklistItem K2WLItem in K2WList)
                {
                    string serialNumber = K2WLItem.SerialNumber;
                    string status = K2WLItem.Status.ToString();
                    //to access more information, drill down into the relevant object contained in the worklist item
                    string Folio = K2WLItem.ProcessInstance.Folio;
                }
            }
        }

        //opening and completing a worklist item samples
        public static void OpenWorklistItemSamples()
        {
            //you must open a K2 connection first. We will wrap the K2 connection into a using statement to ensure that it is closed
            using (Connection K2Conn = new Connection())
            {
                K2Conn.Open("[servername]");
                
                //to open a worklist item you require at least the item's serial number
                //opening a worklist item without any overrides will automatically set the item's status to "Opened" by the connected user.
                WorklistItem K2WListItem = K2Conn.OpenWorklistItem("[serialnumber]");
                

                //once you have opened the item, you can read data from the item or update data in the workflow
                string oldFolio = K2WListItem.ProcessInstance.Folio;
                K2WListItem.ProcessInstance.Folio = "[NewFolio]";
                K2WListItem.ProcessInstance.DataFields["[String DataField Name]"].Value = "[updatedvalue]";

                //to update the process without completing the task, call Update as follows.
                
             
                //to update the worklist item and finish the task, you must call the Action.Execute method
                //the workflow will then continue executing according to the action selected
                //warning: you must select one of the available actiosn for the current task, otherwise an error will be reported
                K2WListItem.Actions["[ActionName]"].Execute();

                //to get the available actions for the task, query the Actions collection
                foreach (SourceCode.Workflow.Client.Action action in K2WListItem.Actions)
                {
                    //do something with the actions. Normally, you would output the available actions into a drop-down list 
                    //or perhaps generate new button controls for each available action
                }

                //if you have want to open a worklist item from a managed user's tasklist, you need to use the OpenManagedWorklistItem method:
                K2WListItem = K2Conn.OpenManagedWorklistItem("[managedUserUsername]", "[serialNumber]");

                //if you are openeing another user's worklist item that was delegated to the current account with Out of Office, you need
                //to use the OpenSharedWorklistItem method
                K2WListItem = K2Conn.OpenSharedWorklistItem("[originalUserName]", "[managedUserName]", "[serialNumber]");
                
                
                //if you do not have the serial number, you can iterate over the worklist to open a worklist item
                //note: returning the worklist and iterating over the list can be an expensive operation
                //for the purposes of this exercise we will check if the worklist item folio is a specific value. A more efficient way would be to set up a 
                //worklist criteria filter to filter the tasks based on the folio we are looking for
                Worklist K2WList = K2Conn.OpenWorklist();
                //once you have retrived the worklist, you can iterate over the worklist items in the worklist
                foreach (WorklistItem K2WLItem in K2WList)
                 {
                     if (K2WLItem.ProcessInstance.Folio == "[somefolio]")
                     {
                         //you must open the worklist item before you can update it or complete it
                         K2WLItem.Open();
                         K2WLItem.ProcessInstance.Folio = "[NewFolio]";
                         K2WLItem.ProcessInstance.DataFields["[SomeDataField]"].Value = "[updatedvalue]";
                         //to update the worklist item and finish it, you must call the Action.Execute method
                         //the workflow will then continue executing accoridng to the action selected
                         K2WLItem.Actions["[ActionName]"].Execute();
                     }
                 }
            }
        }

        //impersonating another user samples
        public static void ImpersonateSamples()
        {
            //we will wrap the connection into a using statement to ensure it is disposed properly
            using (Connection K2Conn = new Connection())
            {
                //you must first establish a connection with the current credentials (or specific credentials) before you can impersonate
                K2Conn.Open("[servername]");
                //once you have connected to K2, you can impersonate another user, PROVIDED that the connected account
                //has the "Impersonate" permission on the K2 environment
                K2Conn.ImpersonateUser("[securityLabelName]:[username]"); //example: K2Conn.ImpersonateUser("K2:domain\username");
                //now that you have impersonated a user, you can perform actiosn on behalf of that user. here we will retrieve that user's worklist
                Worklist K2WL = K2Conn.OpenWorklist();
                //do something with the impersonated user's worklist

                //when you are done impersonating, you can revert to the original user account
                K2Conn.RevertUser();
            }
        }

        //retrieve comments samples
        public static void GetCommentsSamples()
        {
            using (Connection K2Conn = new Connection())
            {
                //dummy values for process instance ID and activity instance ID
                int processInstanceId = 1;
                int activityInstanceDestinationID = 10;

                //retrieve comments for a workflow instance in different ways
                //Get all the Comments using the Process Instance's ID
                IEnumerable<IWorkflowComment> comments = K2Conn.GetComments("[ProcessInstanceID]");
                foreach (IWorkflowComment comment1 in comments)
                {
                    Console.WriteLine(comment1.Message);
                }

                //Get all the Comments using a Worklist Item's SerialNumber
                IEnumerable<IWorkflowComment> comments1 = K2Conn.GetComments("[serialnumber]");
                

                //Get all the Comments using a Process Instance's ID and Activity Instance Destination's ID
                IEnumerable<IWorkflowComment> comments2 = K2Conn.GetComments(processInstanceId, activityInstanceDestinationID);

                //Get all the Comments from the Process Instance's 'Comments' property
                ProcessInstance _processInstance = K2Conn.OpenProcessInstance(processInstanceId);
                IEnumerable<IWorkflowComment> procInstComments = _processInstance.Comments;
                foreach (IWorkflowComment comments3 in procInstComments)
                {
                    Console.WriteLine(comments3.Message);
                }

                //Get all the Comments from the Worklist Item's 'Comments' property
                WorklistItem _worklistItem = K2Conn.OpenWorklistItem("[_serialNo]");
                IEnumerable<IWorkflowComment> WLItemComments = _worklistItem.Comments;
            }
        }

        //add comments samples
        public static void AddCommentsSamples()
        {
            using (Connection K2Conn = new Connection())
            {
                //dummy values for process instance ID and activity instance ID
                int processInstanceId = 1;
                int activityInstanceDestinationID = 10;

                //adding a comment
                //Add a Comment using the Connection class and the Process Instance's ID
                IWorkflowComment comment = K2Conn.AddComment(processInstanceId, "Hello World");

                //Add a Comment using the Connection class and the Worklist Item's SerialNumber
                IWorkflowComment comment2 = K2Conn.AddComment("[_serialNo]", "Hello World");

                //Add a Comment using the Connection class and the Process Instance's ID and Activity Instance Destination's ID
                IWorkflowComment comment3 = K2Conn.AddComment(processInstanceId, activityInstanceDestinationID, "Hello World");

                //Add a Comment using the Process Instance
                ProcessInstance procinst = K2Conn.OpenProcessInstance(processInstanceId);
                IWorkflowComment procInstComment = procinst.AddComment("Hello World");

                //Add a Comment using the WorklistItem
                WorklistItem worklistItem = K2Conn.OpenWorklistItem("[_serialNo]");
                IWorkflowComment WLItemComment = worklistItem.AddComment("Hello World");
            }
        }

        //retrieve attachments samples
        public static void GetAttachmentsSamples()
        {
            using (Connection K2Conn = new Connection())
            {
                //dummy values for process instance ID and activity instance ID
                int processInstanceId = 1;
                int activityInstanceDestinationID = 10;

                //Get all the Attachments using the Process Instance's ID
                //By default this call will always return the attachment's files.
                IEnumerable<IWorkflowAttachment> attachments = K2Conn.GetAttachments(processInstanceId);
                foreach (IWorkflowAttachment attachment1 in attachments)
                {
                    //retrieve the attachment file into an IO stream
                    System.IO.Stream fileContents = attachment1.GetFile();
                }

                //Get all the Attachments using the SerialNumber
                //By default this call will always return the attachment's files.
                IEnumerable<IWorkflowAttachment> attachments2 = K2Conn.GetAttachments("[_serialNo]");

                //Get all the Attachments using the Process Instance's ID and Activity Instance Destination's ID
                //By default this call will always return the attachment's files.
                IEnumerable<IWorkflowAttachment> attachments3 = K2Conn.GetAttachments(processInstanceId, activityInstanceDestinationID);

                //Get all the Attachments from the Process Instance's 'Attachments' property
                ProcessInstance _processInstance = K2Conn.OpenProcessInstance(processInstanceId);
                IEnumerable<IWorkflowAttachment> procInstAttachments = _processInstance.Attachments;

                //Get all the Attachments from the Worklist Item's 'Attachments' property
                WorklistItem _worklistItem = K2Conn.OpenWorklistItem("[_serialNo]");
                IEnumerable<IWorkflowAttachment> WLItemAttachments = _worklistItem.Attachments;

                //Get the Attachment by passing the Attachment's ID
                //By default this call will always return the attachment's file.
                int attachmentId = 1;
                IWorkflowAttachment attachment = K2Conn.GetAttachment(attachmentId);
                //save the attachment.
                SaveFile(attachment.FileName, attachment.GetFile());

                //Get the Attachment by passing the Attachment's ID
                //pass 'false' for the includeFile parameter to only load the file on demand.
                IWorkflowAttachment attachmentNoFile = K2Conn.GetAttachment(attachmentId, false);
             
                //Get all the Attachments using the Process Instance's ID and Activity Instance Destination's ID
                //pass 'false' for the includeFile parameter to only load the file on demand.
                IEnumerable<IWorkflowAttachment> attachments4 = K2Conn.GetAttachments(processInstanceId, activityInstanceDestinationID, false);

                //Iterate through the attachments
                foreach (IAttachment attachment4 in attachments4)
                {
                    //only get specific user's attachments
                    if (attachment4.FQN.Equals(@"K2:Denallix\Bob"))
                    {
                        //use the attachment.GetFile() method to load the attachment's file and cache it.
                        using (System.IO.Stream downloadStream = attachment.GetFile())
                        {
                            SaveFile(attachment.FileName, downloadStream);
                        }
                    }
                }
            }

        }

        //add attachments samples
        public static void AddAttachmentsSamples()
        {
            using (Connection K2Conn = new Connection())
            {
                //dummy values for process instance ID and activity instance ID
                int processInstanceId = 1;
                int activityInstanceDestinationID = 10;
                string fileName = "Report1.pdf";

                //Add an Attachment using the Connection class and the Process Instance's ID
                IWorkflowAttachment _attachment = K2Conn.AddAttachment(processInstanceId, fileName, GetFile("Report1.pdf"));

                //Add an Attachment using the Connection class and the Worklist Item's SerialNumber
                IWorkflowAttachment _attachment1 = K2Conn.AddAttachment("[_serialNo]", fileName, GetFile("Report1.pdf"));

                //Add an Attachment using the Connection class and the Process Instance's ID and Activity Instance Destination's ID
                IWorkflowAttachment _attachment2 = K2Conn.AddAttachment(processInstanceId, activityInstanceDestinationID, fileName, GetFile("Report1.pdf"));

                //Add an Attachment using the Process Instance
                ProcessInstance _processInstance  = K2Conn.OpenProcessInstance(processInstanceId);
                IWorkflowAttachment _attachment3 = _processInstance.AddAttachment(fileName, GetFile("Report1.pdf"));

                //Add an Attachment using the WorklistItem
                WorklistItem _worklistItem = K2Conn.OpenWorklistItem("[_serialNo]");
                IWorkflowAttachment _attachment4 = _worklistItem.AddAttachment(fileName, GetFile("Report1.pdf"));

                //Async exampe, where you create the attachment first and then upload the file 
                //Add an 'empty' attachment
                IWorkflowAttachment _attachment5 = K2Conn.AddAttachment(processInstanceId, fileName, null);
                //now upload the attachment's file.
                //Note. You can only upload the file once and only for an 'empty' attachment.
                //Note. This can be used for async purposes, to create the metadata first and secondly upload the file.
                IAttachment attachmentWithContent = K2Conn.UploadAttachmentContent(_attachment.Id, GetFile("Report1.pdf"));
            }
        }

        #region HelperMethods
        //Save a file's content
        private static void SaveFile(string fileName, System.IO.Stream stream)
        {
            string fullPath = string.Format(@"C:\Temp\{0}", fileName);

            //Check if the file exists. If it does, throw new exception. else save the file.
            if (System.IO.File.Exists(fullPath))
            {
                throw new Exception("File already saved");
            }
            else
            {
                //Create the file
                using (var fileStream = System.IO.File.Create(string.Format(@"C:\Temp\{0}", fileName)))
                {
                    //Note. CopyTo is only available from .Net 4 http://msdn.microsoft.com/en-us/library/dd782932(v=vs.100).aspx
                    stream.CopyTo(fileStream); //Copy the stream to the new file.
                }
            }
        }

        //Get a file's content
        private static System.IO.Stream GetFile(string fileName)
        {
            System.IO.Stream fileStream = null;
            string fullPath = string.Format(@"C:\Temp\{0}", fileName);
            //Check if the file exists at the location.
            if (System.IO.File.Exists(fullPath))
            {
                fileStream = System.IO.File.OpenRead(fullPath);//Get the FileStream
            }

            return fileStream;
        }
        #endregion 
    }
}
