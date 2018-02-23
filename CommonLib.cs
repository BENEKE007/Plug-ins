using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLib
{
    public class Common
    {
        //On Prem
        public Entity RetrieveCrmRecord(IOrganizationService service, string entity, string criteriaFieldName, string criteriaFieldValue, string[] returnFields, bool returnAllFields = false)
        {
            Entity returnValue = null;

            QueryExpression query = new QueryExpression(entity);
            if (returnAllFields)
            {
                query.ColumnSet = new ColumnSet(true);
            }
            else
            {
                query.ColumnSet = new ColumnSet(returnFields);
            }
            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition(new ConditionExpression(criteriaFieldName, ConditionOperator.Equal, criteriaFieldValue));

            EntityCollection ec = service.RetrieveMultiple(query);
            if (ec.Entities.Count > 0)
            {
                returnValue = ec.Entities[0];

            }
            return returnValue;
        }

        public EntityCollection RetrieveCrmRecords(IOrganizationService service, string entity, string criteriaFieldName, string criteriaFieldValue, string[] returnFields, bool returnAllFields = false)
        {
            //AttributeQueryExpression at = new AttributeQueryExpression();
            //at.Criteria = new MetadataFilterExpression();
            //at.Criteria.Conditions.Add(new MetadataConditionExpression("Name", MetadataConditionOperator.Equals, "cebmain_toinvoicequantity"));

            //service.RetrieveMultiple(at);

            QueryExpression query = new QueryExpression(entity);
            if (returnAllFields)
            {
                query.ColumnSet = new ColumnSet(true);
            }
            else
            {
                query.ColumnSet = new ColumnSet(returnFields);
            }
            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition(new ConditionExpression(criteriaFieldName, ConditionOperator.Equal, criteriaFieldValue));

            return service.RetrieveMultiple(query);

        }

        public void ReadAttributeMetaDataForInvoiceLine(IOrganizationService service)
        {
            RetrieveEntityRequest req = new RetrieveEntityRequest();
            req.EntityFilters = EntityFilters.All;
            req.LogicalName = "invoicedetail";

            RetrieveEntityResponse resp = (RetrieveEntityResponse)service.Execute(req);
            foreach (AttributeMetadata a in resp.EntityMetadata.Attributes)
            {
                Console.WriteLine(((AttributeMetadata)(a)).EntityLogicalName + " "
                    + (a.LogicalName) + " " + a.SchemaName);
            }
            Console.ReadLine();

        }

        public RetrieveMetadataChangesResponse GetAttributeDetails(IOrganizationService service)
        {
            MetadataFilterExpression EntityFilter = new MetadataFilterExpression(LogicalOperator.And);
            EntityFilter.Conditions.Add(new MetadataConditionExpression("LogicalName", MetadataConditionOperator.In, "cebmain_toinvoicequantity"));

            MetadataPropertiesExpression EntityProperties = new MetadataPropertiesExpression()
            {
                AllProperties = false
            };
            EntityProperties.PropertyNames.AddRange(new string[] { "Attributes" });

            MetadataConditionExpression optionsetAttributeName = new MetadataConditionExpression("AttributeOf", MetadataConditionOperator.Equals, null);
            MetadataFilterExpression AttributeFilter = new MetadataFilterExpression(LogicalOperator.And);

            AttributeFilter.Conditions.Add(optionsetAttributeName);

            MetadataPropertiesExpression AttributeProperties = new MetadataPropertiesExpression() { AllProperties = false };
            //foreach (string attrProperty in AttributeProperties)
            //{
            //    AttributeProperties.PropertyNames.Add(attrProperty);
            //}

            EntityQueryExpression entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = EntityFilter,
                Properties = EntityProperties,
                AttributeQuery = new AttributeQueryExpression()
                {
                    Properties = AttributeProperties,
                    Criteria = AttributeFilter
                }
            };

            RetrieveMetadataChangesRequest req = new RetrieveMetadataChangesRequest()
            {
                Query = entityQueryExpression,
                //ClientVersionStamp = clientVersionStamp
            };

            return (RetrieveMetadataChangesResponse)service.Execute(req);

        }

        public Entity RetrieveLinkedRecord(IOrganizationService service, string entityName, string linkedToEntityName
            , string linkFromEntityAttributeName, string linkToEntityAttributeName, string guidId)
        {
            QueryExpression query = new QueryExpression(entityName);
            //query.ColumnSet = new ColumnSet("column1", "coumn2");
            // Or retrieve All Columns
            query.ColumnSet = new ColumnSet(true);

            LinkEntity EntityB = new LinkEntity(entityName, linkedToEntityName, linkFromEntityAttributeName, linkToEntityAttributeName, JoinOperator.Inner);
            EntityB.Columns = new ColumnSet(true);
            EntityB.EntityAlias = "EntityB";
            // Can put condition like this to any Linked entity
            // EntityB.LinkCriteria.Conditions.Add(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
            query.LinkEntities.Add(EntityB);

            //// Join Operator can be change if there is chance of Null values in the Lookup. Use Left Outer join
            //LinkEntity EntityC = new LinkEntity("EntityALogicalName", "EntityCLogicalName", "EntityALinkAttributeName", "EntityCLinkAttributeName", JoinOperator.Inner);
            //EntityC.Columns = new ColumnSet("column1", "coumn2");
            //EntityC.Columns = new ColumnSet("column1", "coumn2");
            //EntityC.EntityAlias = "EntityC";
            //query.LinkEntities.Add(EntityC);

            //query.Criteria.Conditions.Add(new ConditionExpression("status", ConditionOperator.Equal, 1));

            return service.Retrieve(linkedToEntityName, new Guid(guidId), new ColumnSet(true));

        }

        public string GetOptionsSetTextForValue(IOrganizationService service, string entityName, string attributeName, int selectedValue)
        {

            RetrieveAttributeRequest retrieveAttributeRequest = new
            RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };
            // Execute the request.
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            // Access the retrieved attribute.
            Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata retrievedPicklistAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata)
            retrieveAttributeResponse.AttributeMetadata;// Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
            string selectedOptionLabel = null;
            foreach (OptionMetadata oMD in optionList)
            {
                if (oMD.Value == selectedValue)
                {
                    selectedOptionLabel = oMD.Label.LocalizedLabels[0].Label.ToString();
                    break;
                }
            }
            return selectedOptionLabel;
        }
    }
}
