public class Session
{
    public Workbook OpenWorkbook(string filePath)
    {
        // Implementation to open a workbook and return a Workbook object
        throw new NotImplementedException();
    }
}

public class Workbook
{
    public Worksheet GetWorksheet(string worksheetName)
    {
        // Implementation to get a Worksheet object from the workbook
        throw new NotImplementedException();
    }

    public void Save()
    {
        // Implementation to save the workbook
        throw new NotImplementedException();
    }

    public void SaveAs(string filePath)
    {
        // Implementation to save the workbook to a new file path
        throw new NotImplementedException();
    }

    public void Close()
    {
        // Implementation to close the workbook
        throw new NotImplementedException();
    }
}

public class Worksheet
{
    public void MakeActive()
    {
        // Implementation to make the worksheet active
        throw new NotImplementedException();
    }

    public void Delete()
    {
        // Implementation to delete the worksheet
        throw new NotImplementedException();
    }

    public void Rename(string newName)
    {
        // Implementation to rename the worksheet
        throw new NotImplementedException();
    }
}
