using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Server.Helpers
{
    public static class QueryableExtensions
    {   //paginacion 150'
        //https://www.youtube.com/watch?v=KP_OFYFOS8M&list=PLuEZQoW9bRnRHpzGBYKWfW01AFB_-Qgn1&index=8
        public static IQueryable<T> Paginate<T>(this IQueryable<T> queryable, PaginationDTO pagination)
        {
            return queryable
                .Skip((pagination.Page - 1) * pagination.RecordsNumber)
                .Take(pagination.RecordsNumber);
        }
    }
}
