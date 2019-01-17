namespace ReadExcel
{
    public interface IModelTypeImportExcel
    {
        bool ImportExcel(string fileName, string uploadBy);
    }
}