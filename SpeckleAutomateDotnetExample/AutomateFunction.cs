using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office.Word;
using Objects;
using Objects.Geometry;
using Speckle.Automate.Sdk;
using Speckle.Core.Api;
using Speckle.Core.Logging;
using Speckle.Core.Models.Extensions;
using Speckle.Core.Transports;
using System.Text.Json;
static class AutomateFunction
{
  public static async Task Run(
    AutomationContext automationContext,
    FunctionInputs functionInputs
  )
  {
    Console.WriteLine("Starting execution");
    _ = typeof(ObjectsKit).Assembly; // INFO: Force objects kit to initialize

    Console.WriteLine("Receiving version");
    var commitObject = await automationContext.ReceiveVersion();

    Console.WriteLine("Received version: " + commitObject);
    var project = automationContext.AutomationRunData.ProjectId;
    var stream = automationContext.SpeckleClient.StreamGet(project);

    string branchData = File.ReadAllText("./SpeckleAutomateDotnetExample/branches.json");
    Branches? branches = JsonSerializer.Deserialize<Branches>(branchData);

    if(branches == null){
      automationContext.MarkRunFailed("The run failed as the branches data failed to be serialised");
      return;
    }
    var element_assignment = await automationContext.SpeckleClient.BranchGet($"{stream.Id}", branches.ElementAssignment);
    var material_assignment = await automationContext.SpeckleClient.BranchGet($"{stream.Id}", branches.Materials);

    if(element_assignment == null)
    { 
      automationContext.MarkRunFailed("The run failed as the element assignment branch does not exist");
      return;
    }
    if(material_assignment == null)
    { 
      automationContext.MarkRunFailed("The run failed as the material assignment branch does not exist");
      return;
    }

    var latest_element_assignment_commit = element_assignment?.commits.items.LastOrDefault();
    var latest_material_assignment_commit = material_assignment?.commits.items.LastOrDefault();

    if(latest_element_assignment_commit == null)
    {
      automationContext.MarkRunFailed("The run failed as there are no commits on the element assignment branch");
      return;
    }

    if(latest_material_assignment_commit == null)
    { 
      automationContext.MarkRunFailed("The run failed as there are no commits on the material assignment branch");
      return;
    }

    var latest_element_assignment = await automationContext.SpeckleClient.CommitGet($"{stream.Id}", $"{latest_element_assignment_commit.id}");
    var latest_material_assignment = await automationContext.SpeckleClient.CommitGet($"{stream.Id}", $"{latest_element_assignment_commit.id}");
    ServerTransport transport = new ServerTransport(automationContext.SpeckleClient.Account, project);

    Speckle.Core.Models.Base? latest_element_assigment_object = await Operations.Receive(latest_element_assignment.id, remoteTransport: transport, disposeTransports: true);
    Speckle.Core.Models.Base? latest_material_assigment_object = await Operations.Receive(latest_material_assignment_commit.id, remoteTransport: transport, disposeTransports: true);
    
    if(latest_element_assigment_object == null){
      automationContext.MarkRunFailed("The run failed as the latest commit on the element assignment branch could not be retrieved");
      return;
    }

    if(latest_material_assigment_object == null){
      automationContext.MarkRunFailed("The run failed as the latest commit on the material assignment branch could not be retrieved");
      return;
    }


    var allObjects = await automationContext.ReceiveVersion();
    Console.WriteLine($"There are: {allObjects.Flatten().Count()} objects");
    automationContext.MarkRunSuccess($"Counted objects");
  }
}
