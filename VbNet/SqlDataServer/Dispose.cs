using System;
namespace SqlDataServer
{ 
public abstract class ReleaseObj : IDisposable
{
    public new void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    protected virtual new void Dispose(bool Disposing)
    {
        if (Disposing)
        {
        }
    }
    ~ReleaseObj()
    {
        Dispose(false);
    }
}
}