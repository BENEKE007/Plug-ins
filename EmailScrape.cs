
using System;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Collections.Generic;
using CommonLib;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Query;

namespace Microsoft.OMI
{
   public class EmailScrape : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {

            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            Entity postEmail = (Entity)context.PostEntityImages["emailPostImage"];

            tracer.Trace("Executed Email Post Image ");

            CommonLib.Common common = new Common();
            Entity email = common.RetrieveCrmRecord(service, "email", "activityid", postEmail.Id.ToString(), new string[] { "" }, true);

            string text = HTMLToText(email.Attributes["description"].ToString());
            string subject = email.Attributes["subject"].ToString();

             tracer.Trace("Email subject " + subject);

            tracer.Trace("Executed Email Post Image " + text);


            int startIndexCN = text.IndexOf("Client Name:");
            int startIndexCP = text.IndexOf("Client Phone:");
            int startIndexCE = text.IndexOf("Client Email:");
            int startIndexSR = text.IndexOf("SR Number :");  //SR Number : mint_scrapedsrnumber
            int startIndexPN = text.IndexOf("Policy Number:"); //Policy Number: mint_scrapedpolicynumber
            int startIndexQN = text.IndexOf("Quote number:");//Quote number: mint_scrapedquotenumber
            int startIndexSV = text.IndexOf("Quote Value:");  //mint_scrapedvalue

            tracer.Trace("startIndexCN " + startIndexCN);
            tracer.Trace("startIndexCP " + startIndexCP);
            tracer.Trace("startIndexCE " + startIndexCE);
            tracer.Trace("startIndexSR " + startIndexSR);
            tracer.Trace("startIndexPN " + startIndexPN);
            tracer.Trace("startIndexQN " + startIndexQN);
            tracer.Trace("startIndexSV " + startIndexSV);

            string nameC ="";
            if (startIndexCN > 0) {  nameC = text.Substring(startIndexCN + 12, startIndexCP - (startIndexCN + 12)); }
            string phoneC = "";
            if (startIndexCN > 0) { phoneC = text.Substring(startIndexCP + 13, startIndexCE - (startIndexCP + 13)); }
            string emailC = "";
            if (startIndexCN > 0) {  emailC = text.Substring(startIndexCE + 13, startIndexSR - (startIndexCE + 13)); }
            string srNumber = "";
            if (startIndexCN > 0) {  srNumber = text.Substring(startIndexSR + 11, startIndexPN - (startIndexSR + 11)); }
            string policyN = "";
            if (startIndexCN > 0) {  policyN = text.Substring(startIndexPN + 14, startIndexQN - (startIndexPN + 14)); }
            string quoteN = "";
            if (startIndexCN > 0) {  quoteN = text.Substring(startIndexQN + 13, startIndexSV - (startIndexQN + 13)); }
            string quoteV = "";
            Money quoteVm = new Money(0);
            if (startIndexCN > 0)
            {
                quoteV = text.Substring(startIndexSV + 13, text.Length - (startIndexSV + 13));  //string s2 = string.Format("{0:#,0.#}", float.Parse(s));
                quoteVm = new Money(decimal.Parse(quoteV));
            }

          

            tracer.Trace("name ");
           // tracer.Trace("name " + nameC); //mint_scrapedclientname
            tracer.Trace("phone " + phoneC);
            tracer.Trace("email " + emailC);
            tracer.Trace("srNumber " + srNumber);

            email.Attributes["mint_scrapedclientname"] = nameC;
            email.Attributes["mint_scrapedsrnumber"] = srNumber;
            email.Attributes["mint_scrapedpolicynumber"] = policyN;
            email.Attributes["mint_scrapedquotenumber"] = quoteN;
            email.Attributes["mint_scrapedclientemail"] = emailC;
            email.Attributes["mint_scrapedclientphone"] = phoneC;
            email.Attributes["mint_scrapedvalue"] = quoteVm;  // Money mny = new Money(decimal.Parse(objCol.txt1));

            service.Update(email);

            string srSubject = "";
            int ind = 0;
            if (subject.Contains("SR"))   //  SR – 1234567 – RFQ
            {
                tracer.Trace("Contains SR ");
                
                 ind = subject.IndexOf("SR");
                tracer.Trace("ind  " + ind.ToString());
                if (ind > 0)
                {
                    tracer.Trace("Contains SR f (ind > 0) ");
                    srSubject = subject.Substring(ind, 18); // text.Substring(startIndexQN + 12, startIndexSV - (startIndexQN + 12));
                }
                tracer.Trace("Contains SR srSubject " + srSubject);
            }

            //Entity regardingCase = common.RetrieveCrmRecord(service, "incident", "mint_srnumber1", srNumber, new string[] { "" }, true);
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "mint_srnumber1";
            condition1.Operator = ConditionOperator.Equal;
            if(ind > 0)
            {
                condition1.Values.Add(srSubject.Trim());
            }
            else
            {
                condition1.Values.Add(srNumber.Trim());
            }
           

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);

            QueryExpression query = new QueryExpression("incident");
            query.ColumnSet.AddColumns("mint_srnumber1", "incidentid");
            query.Criteria.AddFilter(filter1);

            EntityCollection regardingCase = service.RetrieveMultiple(query);
            tracer.Trace("regardingCase Count" + regardingCase.Entities.Count);

            tracer.Trace("regardingCase ");

            if (regardingCase.Entities.Count > 0)
            {
                tracer.Trace("regardingCase 1 ");
                tracer.Trace("regardingCase ID - " + regardingCase[0].Attributes["incidentid"].ToString());

                Guid relatedCase = new Guid(regardingCase[0].Attributes["incidentid"].ToString());
                email.Attributes["regardingobjectid"] = new EntityReference("incident", relatedCase); //(EntityReference)regardingCase.Attributes["incidentid"];
                service.Update(email);

            }

            tracer.Trace("regardingCase ID" );
            tracer.Trace("regardingCase ID " + regardingCase[0].Id.ToString());




        }

        public string HTMLToText(string htmlString)

        {

            string ss = string.Empty;

            Regex regex = new Regex("\\<[^\\>]*\\>");

            ss = regex.Replace(htmlString, String.Empty);
            ss = ss.Replace("&nbsp;", "");

            return ss;// Plain Text as a OUTPUT

        }


    }
}
