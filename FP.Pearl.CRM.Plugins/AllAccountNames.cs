using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Text.RegularExpressions;

namespace FP.Pearl.CRM.Plugins
{
    public class AllAccountNames
    {
        EntityCollection results = new EntityCollection();
        EntityCollection accResults = new EntityCollection();
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, string sName)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select ID, Contacts, Name, Counties, Last Opportunity Assignment, Distribution Rules
            from
                Account a
                    inner join
                Counties b
                    on a.pearl_account_pearl_county = b.pearl_county
                    inner join 
                ZipCodes c
                    on b.pearl_county_pearl_zipcode = c.pearl_zipcode
            where 
            (
	            a.Name = "incoming string"
            )
                */

            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = Account.EntityLogicalName,
                ColumnSet = new ColumnSet("ID", "Contacts", "Name", "Counties", "Last Opportunity Assignment", "Distribution Rules"),
                LinkEntities = 
        {
            new LinkEntity 
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "a.pearl_account_pearl_county",
                LinkFromEntityName = Account.EntityLogicalName,
                LinkToAttributeName = "b.pearl_county",
                LinkToEntityName = pearl_county.EntityLogicalName
            },
            new LinkEntity
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "b.pearl_county_pearl_zipcode",
                LinkFromEntityName = pearl_county.EntityLogicalName,
                LinkToAttributeName = "c.pearl_zipcode",
                LinkToEntityName = pearl_zipcode.EntityLogicalName,
            }
        },
                Criteria =
                {
                    Filters = 
            {
                new FilterExpression
                {
                    Conditions = 
                    {
                        new ConditionExpression("Name", ConditionOperator.Equal, sName)
                    },
                }
            }
                }
            };

            entityCollection = service.RetrieveMultiple(query).Entities;

            return entityCollection;
        }
    }
}
