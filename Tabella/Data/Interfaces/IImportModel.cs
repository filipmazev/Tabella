namespace Tabella.Data.Interfaces;

/// <summary>
/// Interface for import models, providing a method to convert to an app item.
/// </summary>
/// <typeparam name="TAppItem"></typeparam>
public interface IImportModel<out TAppItem> where TAppItem : class
{
    /// <summary>
    /// Method that converts the import model to an app item.
    /// </summary>
    /// <returns></returns>
    TAppItem ToAppItem();
}