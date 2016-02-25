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
    public class PearlAnyTwoKeyValue
    {
        EntityCollection results = new EntityCollection();
        EntityCollection accResults = new EntityCollection();
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, int iValue)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select pearl_text1
            from
                pearl_anytwokey
            where 
            (
	            pearl_action = input value
            )
            */

            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = pearl_anytwokey.EntityLogicalName,
                ColumnSet = new ColumnSet("pearl_text1"),
                LinkEntities =
                {
                },
                Criteria =
                {
                    Filters = 
            {
                new FilterExpression
                {
                    Conditions = 
                    {
                        new ConditionExpression("pearl_action", ConditionOperator.Equal, iValue)
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
