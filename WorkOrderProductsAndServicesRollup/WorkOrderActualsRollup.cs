using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace WorkOrderProductsAndServicesRollup
{
    public class WorkOrderActualsRollup : IPlugin
    {
        private IOrganizationService service;
        private ITracingService trace;
        private string primaryEntity;
        private Guid workOrderId;

        public void Execute(IServiceProvider serviceProvider)
        {
            trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Handle Create/Update
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    if (target.LogicalName != "msdyn_workorderservice" &&
                        target.LogicalName != "msdyn_workorderproduct")
                        return;

                    primaryEntity = target.LogicalName;

                    if (context.PostEntityImages.Contains("Image"))
                    {
                        var img = context.PostEntityImages["Image"];

                        if (img.Contains("msdyn_workorder") &&
                            img.Contains("msdyn_linestatus") &&
                            (img.Contains("msdyn_totalamount") || img.Contains("msdyn_totalcost")))
                        {
                            workOrderId = img.GetAttributeValue<EntityReference>("msdyn_workorder").Id;

                            var totalAmount = img.GetAttributeValue<Money>("msdyn_totalamount");
                            var totalCost = img.GetAttributeValue<Money>("msdyn_totalcost");

                            if (totalAmount != null || totalCost != null)
                                CalculateTotals(workOrderId);
                        }
                    }
                }
                else
                {
                    // Handle Delete
                    if (context.MessageName == "Delete" && context.PreEntityImages.Contains("Image"))
                    {
                        var img = context.PreEntityImages["Image"];

                        primaryEntity = ((EntityReference)context.InputParameters["Target"]).LogicalName;

                        if (img.Contains("msdyn_workorder") &&
                            img.Contains("msdyn_linestatus") &&
                            (img.Contains("msdyn_totalamount") || img.Contains("msdyn_totalcost")))
                        {
                            var lineStatus = img.GetAttributeValue<OptionSetValue>("msdyn_linestatus");
                            var amount = img.GetAttributeValue<Money>("msdyn_totalamount");
                            var cost = img.GetAttributeValue<Money>("msdyn_totalcost");

                            workOrderId = img.GetAttributeValue<EntityReference>("msdyn_workorder").Id;

                            // Only recalc if Used line deleted
                            if (lineStatus.Value == 690970001 && (amount != null || cost != null))
                                CalculateTotals(workOrderId);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                trace.Trace("WorkOrderActualsRollup Error: " + ex.ToString());
                throw;
            }
        }


        private void CalculateTotals(Guid workOrderId)
        {
            decimal totalAmount = 0m;
            decimal totalCost = 0m;
            decimal upsoldProductsAmount = 0m;
            decimal upsoldServicesAmount = 0m;

            QueryExpression query = new QueryExpression(primaryEntity);

            if (primaryEntity == "msdyn_workorderservice")
            {
                query.ColumnSet = new ColumnSet(
                    "msdyn_totalamount",
                    "msdyn_totalcost",
                    "msdyn_workorder",
                    "msdyn_linestatus",
                    "statuscode",
                    "bolt_upsoldservice" // NEW FIELD
                );
            }
            else // msdyn_workorderproduct
            {
                query.ColumnSet = new ColumnSet(
                    "msdyn_totalamount",
                    "msdyn_totalcost",
                    "msdyn_workorder",
                    "msdyn_linestatus",
                    "statuscode",
                    "tb_upsoldproduct"
                );
            }

            query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
            query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, workOrderId);

            var lines = service.RetrieveMultiple(query);

            foreach (var line in lines.Entities)
            {
                var status = line.GetAttributeValue<OptionSetValue>("msdyn_linestatus");

                if (status != null && status.Value == 690970001) // Used
                {
                    if (line.Contains("msdyn_totalamount"))
                        totalAmount += line.GetAttributeValue<Money>("msdyn_totalamount").Value;

                    if (line.Contains("msdyn_totalcost"))
                        totalCost += line.GetAttributeValue<Money>("msdyn_totalcost").Value;

                    // NEW — UPSOLD SERVICE LOGIC
                    if (primaryEntity == "msdyn_workorderservice" &&
                        line.GetAttributeValue<bool>("bolt_upsoldservice") == true &&
                        line.Contains("msdyn_totalamount"))
                    {
                        upsoldServicesAmount += line.GetAttributeValue<Money>("msdyn_totalamount").Value;
                    }

                    // EXISTING — UPSOLD PRODUCT LOGIC
                    if (primaryEntity == "msdyn_workorderproduct" &&
                        line.GetAttributeValue<bool>("tb_upsoldproduct") == true &&
                        line.Contains("msdyn_totalamount"))
                    {
                        upsoldProductsAmount += line.GetAttributeValue<Money>("msdyn_totalamount").Value;
                    }
                }
            }

            // Update Work Order
            Entity wo = new Entity("msdyn_workorder")
            {
                Id = workOrderId
            };

            if (primaryEntity == "msdyn_workorderservice")
            {
                wo["tb_servicetotalcost"] = totalCost;
                wo["tb_servicetotalprice"] = totalAmount;
                wo["bolt_upsoldservicestotal"] = upsoldServicesAmount; // NEW ROLLUP FIELD
            }
            else
            {
                wo["tb_productstotalcost"] = totalCost;
                wo["tb_productstotalprice"] = totalAmount;
                wo["tb_upsoldproductstotal"] = upsoldProductsAmount;
            }

            service.Update(wo);
        }
    }
}
