//MIT License
//
//Copyright (c) 2024 Microsoft - Jose Batista-Neto
//
//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:
//
//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

using System.Data;
using System.Linq.Expressions;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;

namespace GraphExportAPIforMicrosoftTeamsSample.DB;

// This is a helper class for the Entity Framework Functions
// It is used to perform CRUD operations on the database
// all CRUS Operations are done using ADO.net I did this to have a better control of the transactions and to avoid the overhead of the Entity Framework
internal class DbEntity<TEntity> where TEntity : class
{
    private readonly DbContext context;
    private readonly DbSet<TEntity> dbSet;

    public DbEntity(DbContext context)
    {
        this.context = context;
        dbSet = context.Set<TEntity>();
    }

    public async Task<TEntity?> GetFirst()
    {
        return await dbSet.FirstOrDefaultAsync();
    }

    public async Task<IEnumerable<TEntity>> GetAllAsync()
    {
        return await dbSet.ToListAsync();
    }

    public TResult? FindFirst<TResult>(Expression<Func<TEntity, bool>> predicate, Expression<Func<TEntity, TResult>> selector)
    {
        return dbSet.Where(predicate).Select(selector).FirstOrDefault();
    }

    public async Task<IEnumerable<TResult>> FindAll<TResult>(Expression<Func<TEntity, bool>> predicate, Expression<Func<TEntity, TResult>> selector)
    {
        return await dbSet.Where(predicate).Select(selector).ToListAsync();
    }

    public async Task<TEntity?> GetByKeyAsync(object[] Key)
    {
        return await dbSet.FindAsync(Key);
    }

    public async Task AddAsync(TEntity entity)
    {
        await dbSet.AddAsync(entity);
        await context.SaveChangesAsync();
    }

    public async Task UpdateAsync(TEntity entity)
    {
        dbSet.Attach(entity);
        dbSet.Entry(entity).State = EntityState.Modified;
        await context.SaveChangesAsync();
    }

    public async Task DeleteAsync(object[] key)
    {
        TEntity? entityToDelete = await dbSet.FindAsync(key);
        if (entityToDelete != null)
        {
            dbSet.Remove(entityToDelete);
            await context.SaveChangesAsync();
        }
    }

    public async Task DeleteAsyncTransactional(object[] key)
    {
        using (IDbContextTransaction transaction = await context.Database.BeginTransactionAsync(System.Data.IsolationLevel.ReadCommitted))
        {
            try
            {
                TEntity? entityToDelete = await dbSet.FindAsync(key);
                if (entityToDelete != null)
                {
                    dbSet.Remove(entityToDelete);
                    await context.SaveChangesAsync();
                }
                await transaction.CommitAsync();
            }
            catch
            {
                await transaction.RollbackAsync();
                throw; // Re-throw the exception to be handled by the caller
            }
        }
    }

    // Expose transaction-related methods
    public async Task<IDbContextTransaction> BeginTransactionAsync(IsolationLevel isolationLevel)
    {
        return await context.Database.BeginTransactionAsync(isolationLevel);
    }

    // Call this method after performing all operations to save changes
    public async Task SaveChangesAsync()
    {
        await context.SaveChangesAsync();
    }
}
