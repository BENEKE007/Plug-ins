
using System;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Collections.Generic;
using CommonLib;
using System.Linq;

namespace Microsoft.OMI
{
	public class SrNumberPlugin: IPlugin
	{

        public void Execute(IServiceProvider serviceProvider)
        {

            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);          

            Entity postCase = (Entity)context.PostEntityImages["casePostImage"];

            tracer.Trace("Executed Case Post Image " + postCase.LogicalName);

            CommonLib.Common common = new Common();
            Entity incidentCase = common.RetrieveCrmRecord(service, "incident", "incidentid", postCase.Id.ToString(), new string[] { "" }, true);
            // subjectid

            Guid value = Guid.Empty;
           
                tracer.Trace("subjectid != null");
                value = ((EntityReference)(postCase.Attributes["subjectid"])).Id;
          
            
            tracer.Trace("incidentCase Guid value" + value);

            Entity subjectRecord = common.RetrieveCrmRecord(service, "subject", "subjectid", value.ToString(), new string[] { "" }, true);
           

            string text = subjectRecord.Attributes["title"].ToString();
            tracer.Trace("incidentCase subjectRecord text" + text);

            int startIndex = text.IndexOf("(");
            string subject = text.Substring(startIndex +1 , 3);

           // tracer.Trace("subjectText - " + subject);

            if (incidentCase != null)
            {
                tracer.Trace("incidentCase");

                if (incidentCase.Attributes["mint_srnumber2"] == null)
                {
                    tracer.Trace("mint_srnumber2 == null");
                    CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest();
                    createAttributeRequest.EntityName = "incident";

                    var autoNumberAttributeMetadata = new StringAttributeMetadata()
                    {
                        AutoNumberFormat = "SR - {SEQNUM:7} - ",
                        SchemaName = "mint_srnumber2",
                        MaxLength = 50,
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.ApplicationRequired),
                        DisplayName = new Label("Auto Number 2", 1033),
                        Description = new Label("Auto number field through SDK", 1033)
                    };

                    createAttributeRequest.Attribute = autoNumberAttributeMetadata;
                    var response = service.Execute(createAttributeRequest);

                }
                else
                {
                    tracer.Trace("mint_srnumber2 != null");



                   // if (incidentCase.Attributes["mint_srnumber"].ToString() == null)
                   // {
                   
                        incidentCase.Attributes["mint_srnumber2"] = incidentCase.Attributes["mint_srnumber2"] + subject;                     
                        incidentCase.Attributes["mint_srnumber"] = incidentCase.Attributes["mint_srnumber2"];
                    service.Update(incidentCase);
                    tracer.Trace("mint_srnumber -> updated");
                    
                    

                   // }


                     incidentCase.Attributes["mint_srnumber"] = incidentCase.Attributes["mint_srnumber2"];
                    //incidentCase.Attributes["mint_srnumber"] = incidentCase.Attributes["mint_srnumber1"];
                    service.Update(incidentCase);
                    tracer.Trace("UpdateAttributeRequest 3");


                }

            }

            
            tracer.Trace("incidentCase1");


            
        }

        }
}

