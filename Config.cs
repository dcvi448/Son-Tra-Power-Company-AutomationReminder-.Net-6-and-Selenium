public class Alert
{
    public string fromDay { get; set; }
    public string toDay { get; set; }
    public List<string> users { get; set; }
    public Message message { get; set; }
}

public class DeviceName
{
    public string id { get; set; }
    public string choose { get; set; }
    public string isUse { get; set; }
}

public class DeviceType
{
    public string id { get; set; }
    public string choose { get; set; }
    public string isUse { get; set; }
}

public class Download
{
    public string folderDir { get; set; }
}

public class Eoffice
{
    public string token { get; set; }
    public string apiUrl { get; set; }
}

public class ExportXLSX
{
    public string id { get; set; }
}

public class GetAllDevice
{
    public string id { get; set; }
}

public class LevelA
{
    public string id { get; set; }
    public string choose { get; set; }
    public string isUse { get; set; }
}

public class Login
{
    public string id { get; set; }
}

public class Message
{
    public string header { get; set; }
    public string content { get; set; }
    public string footer { get; set; }
}

public class Password
{
    public string id { get; set; }
    public string pass { get; set; }
}

public class Root
{
    public Username username { get; set; }
    public Password password { get; set; }
    public Login login { get; set; }
    public Unit unit { get; set; }
    public Status status { get; set; }
    public UnitName unitName { get; set; }
    public DeviceType deviceType { get; set; }
    public DeviceName deviceName { get; set; }
    public LevelA levelA { get; set; }
    public Team team { get; set; }
    public ToDate toDate { get; set; }
    public GetAllDevice getAllDevice { get; set; }
    public ExportXLSX exportXLSX { get; set; }
    public Download download { get; set; }
    public Eoffice eoffice { get; set; }
    public Alert alert { get; set; }
}

public class Status
{
    public string id { get; set; }
    public Value value { get; set; }
    public string choose { get; set; }
}

public class Team
{
    public string id { get; set; }
    public string choose { get; set; }
    public string isUse { get; set; }
}

public class ToDate
{
    public string id { get; set; }
    public string choose { get; set; }
    public string isUse { get; set; }
}

public class Unit
{
    public string includeAll { get; set; }
    public string id { get; set; }
}

public class UnitName
{
    public string id { get; set; }
    public string choose { get; set; }
    public string isUse { get; set; }
}

public class Username
{
    public string id { get; set; }
    public string user { get; set; }
}

public class Value
{
    public string amostExpire { get; set; }
    public string expire { get; set; }
    public string damage { get; set; }
    public string notYetVerify { get; set; }
}

