using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Extending.Dyn365.v9.CustomWorkflowActivity
{
    /// <summary>
    /// Send birthday greeting to all relevant contacts 
    /// </summary>
    public sealed class SendBirthdayGreeting : CodeActivity
    {
        #region members

        ITracingService tracingService = null;

        #endregion

        #region Custom Workflow Activity Parameters

        [RequiredArgument]
        [Input("Greeting email Title")]
        public InArgument<string> greetingEmailTitle { get; set; }

        [RequiredArgument]
        [Input("Greeting email text")]
        public InArgument<string> greetingEmailtext { get; set; }

        #endregion 

        protected override void Execute(CodeActivityContext context)
        {
            //extract input parameters 
            string greetingTitle = this.greetingEmailTitle.Get(context);
            string greetingText = this.greetingEmailtext.Get(context);

            //initiate organization service proxy
            IWorkflowContext workflowcontext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowcontext.InitiatingUserId);

            //initiate tracing service
            tracingService = context.GetExtension<ITracingService>();
          
            tracingService.Trace("Execution start");

            EntityCollection birthdayContacts = RetrieveBirthdayContacts(service);

            AttachEmailActivity(birthdayContacts, workflowcontext.InitiatingUserId, greetingTitle, greetingText, service);

            tracingService.Trace("Execution end");
        }

        /// <summary>
        /// returns all contact with birthdate that occurs today 
        /// </summary>
        /// <param name="service"></param>
        /// <returns></returns>
        private EntityCollection RetrieveBirthdayContacts(IOrganizationService service)
        {
            EntityCollection result = new EntityCollection();

            string query = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='contact'>
                                <attribute name='fullname' />
                                <attribute name='birthdate' />
                                <filter type='and'>
                                  <condition attribute='birthdate' operator='not-null' />
                                </filter>
                              </entity>
                            </fetch>";

            //retrieve all contacts with birthdate value  
            try
            {
                //perform query
                EntityCollection queryResult = service.RetrieveMultiple(new FetchExpression(query));

                //for each contact, test day and month portions of birthdate 
                foreach (Entity contact in queryResult.Entities)
                {
                    DateTime birthdate = DateTime.Parse(contact["birthdate"].ToString());

                    if ((birthdate.Month == DateTime.Now.Month) && (birthdate.Day == DateTime.Now.Day))
                    {
                        result.Entities.Add(contact);
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace(string.Format("An error occured. Message: {0}\nTrace: {1}", ex.Message, ex.StackTrace));
            }

            return result;
        }

        
        /// <summary>
        /// Attach and send email activity 
        /// </summary>
        /// <param name="contacts"></param>
        private void AttachEmailActivity(
            EntityCollection contacts, Guid currentUserId, string greetingTitle, string greetingText, 
            IOrganizationService service)
        {
            foreach(Entity contact in contacts.Entities)
            {
                //define email activity 
                // define the 'from:' activity party for the email
                EntityReference fromMember = new EntityReference("systemuser", currentUserId);
                Entity fromActivityParty = new Entity("activityparty");
                fromActivityParty["partyid"] = fromMember;

                // define the 'to:' activity party for the email
                EntityReference toMember = new EntityReference("contact", contact.Id);
                Entity toActivityParty = new Entity("activityparty");
                toActivityParty["partyid"] = toMember;

                // Create an e-mail message.
                Entity email = new Entity("email");
                email["from"] = new Entity[] { fromActivityParty }; ;
                email["to"] = new Entity[] { toActivityParty };
                email["subject"] = greetingTitle;
                email["description"] = greetingText;
                email["directioncode"] = true;
                email["regardingobjectid"] = toMember;

                //create email record 
                Guid emailId =  service.Create(email);

                // define SendEmail request.
                SendEmailRequest sendEmailreq = new SendEmailRequest
                {
                    EmailId = emailId,
                    TrackingToken = "",
                    IssueSend = true
                };
           
                //send email 
                SendEmailResponse sendEmailresponse = (SendEmailResponse)service.Execute(sendEmailreq);
            }
        }
    }
}

