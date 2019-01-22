namespace ReadExcel
{
    public interface IModelTypeImportExcel
    {
        /// <summary>
        /// Resprent's import excel file.
        /// </summary>
        /// <param name="fileName">
        /// Full path filename.
        /// </param>
        /// <param name="uploadBy">
        /// Name of uploader.
        /// </param>
        /// <returns></returns>
        bool ImportExcel(string fileName, string uploadBy);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        bool IsPreverifyExcel(UploadFileImportModel model);
    }
}