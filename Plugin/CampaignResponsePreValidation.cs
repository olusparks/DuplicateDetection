using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using BDO.Plugin.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO.Plugin
{
    public class CampaignResponsePreValidation : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext pluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(pluginExecutionContext.UserId);
            ITracingService trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                trace.Trace("Begin Operation");

                if (pluginExecutionContext.InputParameters["Target"] != null)
                {
                    Entity target = (Entity)pluginExecutionContext.InputParameters["Target"];
                    trace.Trace($"Inside target: {target.LogicalName}");

                    if (pluginExecutionContext.MessageName.Equals(PluginMessage.Create, StringComparison.OrdinalIgnoreCase))
                    {
                        trace.Trace("Inside create");

                        string firstName = target.GetAttributeValue<string>(CampaignResponseAttributes.Firstname);
                        string lastName = target.GetAttributeValue<string>(CampaignResponseAttributes.Lastname);
                        string email = target.GetAttributeValue<string>(CampaignResponseAttributes.Email);
                        string companyName = target.GetAttributeValue<string>(CampaignResponseAttributes.Companyname);
                        EntityReference relatedCampaign = target.GetAttributeValue<EntityReference>(CampaignResponseAttributes.Regarding);

                        trace.Trace($"Email: {email}; Company: {companyName}");

                        if (relatedCampaign != null)
                        {
                            //Retrieve all campaign responses related to regarding
                            QueryExpression qeCampaignResponse = new QueryExpression(target.LogicalName);
                            qeCampaignResponse.ColumnSet = new ColumnSet(new string[] { CampaignResponseAttributes.UniqueId });
                            qeCampaignResponse.Criteria.AddCondition(CampaignResponseAttributes.Regarding, ConditionOperator.Equal, relatedCampaign.Id);

                            if (!string.IsNullOrEmpty(email) || !string.IsNullOrEmpty(companyName))
                            {
                                FilterExpression childFilter = new FilterExpression(LogicalOperator.Or);
                                if (!string.IsNullOrEmpty(email))
                                {
                                    childFilter.AddCondition(new ConditionExpression(CampaignResponseAttributes.Email, ConditionOperator.Equal, email));
                                }
                                if (!string.IsNullOrEmpty(companyName))
                                {
                                    childFilter.AddCondition(new ConditionExpression(CampaignResponseAttributes.Companyname, ConditionOperator.Equal, companyName));
                                }
                                qeCampaignResponse.Criteria.AddFilter(childFilter);
                            }
                            qeCampaignResponse.NoLock = true;

                            var campaignResponseCollection = service.RetrieveMultiple(qeCampaignResponse);
                            if (campaignResponseCollection != null && campaignResponseCollection.Entities.Count > 0)
                            {
                                trace.Trace($"Response count {campaignResponseCollection.Entities.Count}");

                                throw new InvalidPluginExecutionException("Duplicate record detected.");
                            }
                        }
                    }
                }

                trace.Trace("End Operation");
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw new InvalidPluginExecutionException(OperationStatus.Failed, ex.Message);
            }
            catch (Exception ex)
            {
                trace?.Trace("An error occurred executing BDO.Plugin: {0}", ex.ToString());
                throw new InvalidPluginExecutionException($"An error occurred executing {nameof(CampaignResponsePreValidation)}.", ex);
            }
        }
    }
}
