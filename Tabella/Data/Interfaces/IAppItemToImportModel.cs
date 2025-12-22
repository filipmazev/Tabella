namespace Tabella.Data.Interfaces;

/// <summary>
/// Interface for app items, providing a method to convert to an import model.
/// </summary>
/// <typeparam name="TImportModel"></typeparam>
public interface IAppItemToImportModel<out TImportModel> where TImportModel : IImportModel<IAppItemToImportModel<TImportModel>>
{
    /// <summary>
    /// Method that converts the app item to an import model.
    /// </summary>
    /// <returns></returns>
    TImportModel ToImportModel();
}