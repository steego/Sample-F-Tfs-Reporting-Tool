module TFSClient

open System
open System.Runtime.Serialization

let tfsAddr = "http://tfs.aiscorp.com:8080/tfs"
let tfsUri = new Uri(tfsAddr)
  

// a series of classes used to hold results from TFS RESTful API queries
[<DataContract>] 
type queryData = 
    { 
        [<DataMember>] mutable query : string 
    }
[<DataContract>] 
type workItemResults = 
    { 
        [<DataMember>] mutable id : string; 
        [<DataMember>] mutable url : string 
    }
[<DataContract>] 
type queryAttributes = 
    { 
        [<DataMember>] mutable referenceName : string; 
        [<DataMember>] mutable name : string; 
        [<DataMember>] mutable url : string 
    }

[<DataContract>]
type queryResults =
    {
        [<DataMember>] mutable queryType : string ;
        [<DataMember>] mutable queryResultType : string;
        [<DataMember>] mutable asOf : string;
        [<DataMember>] mutable columns : ResizeArray<queryAttributes>;
        [<DataMember>] mutable workItems : ResizeArray<workItemResults>;
    }

[<DataContract>] 
type startEndDates =
    {
        [<DataMember>] mutable startDate : string; 
        [<DataMember>] mutable finishDate : string; 
    }

[<DataContract>] 
type iterationData = 
    { 
        [<DataMember>] mutable id : string; 
        [<DataMember>] mutable name : string; 
        [<DataMember>] mutable path : string; 
        [<DataMember>] mutable attributes : startEndDates; 
        [<DataMember>] mutable url : string; 
    }

[<DataContract>]
type iterationResults = 
    {
        [<DataMember>] mutable count : int;
        [<DataMember>] mutable value : ResizeArray<iterationData>
    }

[<DataContract>]
type teamMember = 
    {
        [<DataMember>] mutable id : string;
        [<DataMember>] mutable displayName : string;
        [<DataMember>] mutable uniqueName : string;
        [<DataMember>] mutable url : string;
        [<DataMember>] mutable imageUrl : string;
    }

[<DataContract>]
type activity = 
    {
        [<DataMember>] mutable capacityPerDay : int;
        [<DataMember>] mutable name : string;
    }

[<DataContract>]
type dayOff = 
    {
        [<DataMember>] mutable start : string;
        [<DataMember>] mutable ``end`` : string;
    }

[<DataContract>]
type capacityData = 
    {
        [<DataMember>] mutable teamMember : teamMember;
        [<DataMember>] mutable activities : ResizeArray<activity>;
        [<DataMember>] mutable daysOff : ResizeArray<dayOff>;
        [<DataMember>] mutable url : string;
    }

[<DataContract>]
type capacityResults = 
    {
        [<DataMember>] mutable count : int;
        [<DataMember>] mutable value : ResizeArray<capacityData>
    }

open Microsoft.TeamFoundation

//let tpc = Client.TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(tfsAddr + "/DefaultCollection"));