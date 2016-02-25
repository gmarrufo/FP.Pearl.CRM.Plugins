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
    public class RentalLineByAccount
    {
        EntityCollection results = new EntityCollection();
        EntityCollection accResults = new EntityCollection();
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, string sAccountName)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select ID, new_dealcode, new_GenProdPostingGroup, new_DocumentNo, new_LineNo
            from
                Rental_Line a
                    inner join
                Rental_Header b
                    on a.new_rentalhead_to_rentallineid = b.new_rentalheaderid
                    inner join 
                Account c
                    on b.new_SelltoCustomerNo = c.new_stccustomerno
            where 
            (
	            c.new_AccountName = "incoming string" and
	            a.new_DocumentNo is not null and
	            a.new_LineNo <> 0 and
	            a.new_GenProdPostingGroup = "meter" and
	            (a.new_dealcode not like 'D%' or a.new_dealcode not like 'd%' or a.new_dealcode not like 'C%' or a.new_dealcode not like 'c%')
            )
            */

            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = new_rentalline.EntityLogicalName,
                ColumnSet = new ColumnSet("ID", "new_dealcode", "new_GenProdPostingGroup", "new_DocumentNo", "new_LineNo"),
                LinkEntities = 
        {
            new LinkEntity
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "new_rentalhead_to_rentallineid",
                LinkFromEntityName = new_rentalline.EntityLogicalName,
                LinkToAttributeName = "new_rentalheaderid",
                LinkToEntityName = new_rentalheader.EntityLogicalName
            },
            new LinkEntity
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "new_SelltoCustomerNo",
                LinkFromEntityName = new_rentalheader.EntityLogicalName,
                LinkToAttributeName = "new_stccustomerno",
                LinkToEntityName = Account.EntityLogicalName,
            }
        },
                Criteria =
                {
                    Filters = 
            {
                new FilterExpression
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions = 
                    {
                        new ConditionExpression("c.new_AccountName", ConditionOperator.Equal, sAccountName),
                        new ConditionExpression("a.new_DocumentNo", ConditionOperator.NotNull),
                        new ConditionExpression("a.new_LineNo", ConditionOperator.NotEqual, 0),
                        new ConditionExpression("a.new_GenProdPostingGroup", ConditionOperator.Equal, "meter")
                    },
                },
                new FilterExpression
                {
                    FilterOperator = LogicalOperator.Or,
                    Conditions = 
                    {
                        new ConditionExpression("a.new_dealcode", ConditionOperator.NotLike, "D%"),
                        new ConditionExpression("a.new_dealcode", ConditionOperator.NotLike, "d%"),
                        new ConditionExpression("a.new_dealcode", ConditionOperator.NotLike, "C%"),
                        new ConditionExpression("a.new_dealcode", ConditionOperator.NotLike, "c%")
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
