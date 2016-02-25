// <copyright file="InsertLeadProcess.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author>GMC</author>
// <date>11/1/2015 12:23:59 PM</date>
// <summary>Implements the Insert Lead Process to determine qualification, non qualification among others to the raw lead prior to assignment.</summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Text.RegularExpressions;

namespace FP.Pearl.CRM.Plugins
{
    public class InsertLeadProcess : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                // Lead Entity coming in
                Entity entity = (Entity)context.InputParameters["Target"];

                // Instantiate Organization Service Interfaces
                IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

                try
                {
                    #region STEP -1 Qualified & Not Qualified

                    // Determine Qualified - depending on the pearl_autoqualified value

                    /* QUALIFIED */

                    // Obtain Raw Lead Object
                    if (context.InputParameters.Contains("pearl_autoqualified"))
                    {
                        /*  Check for the pearl_autoqualified value */
                        if (entity.Attributes["pearl_autoqualified"].Equals("true"))
                        {
                            /* Qualify the Incoming Lead */
                            entity.Attributes["statecode"] = new OptionSetValue(1); // Qualified
                            entity.Attributes["statuscode"] = "Qualified";
                        }
                    }

                    #endregion
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the InsertLeadProcess plug-in.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("InsertLeadProcess: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}