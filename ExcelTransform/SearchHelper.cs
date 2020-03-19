using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTransform
{
    public class SearchHelper
    {
        /// <summary>
        /// 通过查询条件生成 lambda 表达式
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="filterConditionList">查询条件</param>
        /// <returns></returns>
        public static Expression<Func<T, bool>> GetFilterExpression<T>(List<SearchFilter> filterConditionList)
        {
            Expression<Func<T, bool>> condition = null;
            try
            {
                if (filterConditionList != null && filterConditionList.Count > 0)
                {
                    foreach (SearchFilter filterCondition in filterConditionList)
                    {
                        Expression<Func<T, bool>> tempCondition = CreateLambda<T>(filterCondition);
                        if (condition == null)
                        {
                            condition = tempCondition;
                        }
                        else
                        {
                            if ("OR".Equals(filterCondition.Relation))
                            {
                                condition = condition.Or(tempCondition);
                            }
                            else
                            {
                                condition = condition.And(tempCondition);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return condition;
        }

        public static Expression<Func<T, bool>> CreateLambda<T>(SearchFilter filterCondition)
        {
            var parameter = Expression.Parameter(typeof(T), "p");//创建参数i
            MemberExpression member = Expression.PropertyOrField(parameter, filterCondition.Property);
            var constant = Expression.Constant(filterCondition.SearchValue);//创建常数

            #region ConstantExpression 类型适配

            if (member.Type == typeof(DateTime))
            {
                var dt = DateTime.Parse(filterCondition.SearchValue);
                constant = Expression.Constant(dt);
            }
            else if (member.Type == typeof(int))
            {
                var i = int.Parse(filterCondition.SearchValue);
                constant = Expression.Constant(i);
            }
            else if (member.Type == typeof(double))
            {
                var d = double.Parse(filterCondition.SearchValue);
                constant = Expression.Constant(d);
            }
            else if (member.Type == typeof(Guid))
            {
                var d = Guid.Parse(filterCondition.SearchValue);
                constant = Expression.Constant(d);
            }

            #endregion ConstantExpression 类型适配

            if ("contains".Equals(filterCondition.Condition))
            {
                return GetExpressionWithMethod<T>("Contains", filterCondition);
            }
            else if ("==".Equals(filterCondition.Condition))
            {
                return Expression.Lambda<Func<T, bool>>(Expression.Equal(member, constant), parameter);
            }
            else if ("!=".Equals(filterCondition.Condition))
            {
                return Expression.Lambda<Func<T, bool>>(Expression.NotEqual(member, constant), parameter);
            }
            else if (">".Equals(filterCondition.Condition))
            {
                return Expression.Lambda<Func<T, bool>>(Expression.GreaterThan(member, constant), parameter);
            }
            else if ("<".Equals(filterCondition.Condition))
            {
                return Expression.Lambda<Func<T, bool>>(Expression.LessThan(member, constant), parameter);
            }
            else if (">=".Equals(filterCondition.Condition))
            {
                return Expression.Lambda<Func<T, bool>>(Expression.GreaterThanOrEqual(member, constant), parameter);
            }
            else if ("<=".Equals(filterCondition.Condition))
            {
                return Expression.Lambda<Func<T, bool>>(Expression.LessThanOrEqual(member, constant), parameter);
            }
            else
            {
                return null;
            }
        }

        public static Expression<Func<T, bool>> GetExpressionWithMethod<T>(string methodName, SearchFilter filterCondition)
        {
            ParameterExpression parameterExpression = Expression.Parameter(typeof(T), "p");
            MethodCallExpression methodExpression = GetMethodExpression(methodName, filterCondition.Property, filterCondition.SearchValue, parameterExpression);
            return Expression.Lambda<Func<T, bool>>(methodExpression, parameterExpression);
        }

        public static Expression<Func<T, bool>> GetExpressionWithoutMethod<T>(string methodName, SearchFilter filterCondition)
        {
            ParameterExpression parameterExpression = Expression.Parameter(typeof(T), "p");
            MethodCallExpression methodExpression = GetMethodExpression(methodName, filterCondition.Property, filterCondition.SearchValue, parameterExpression);
            var notMethodExpression = Expression.Not(methodExpression);
            return Expression.Lambda<Func<T, bool>>(notMethodExpression, parameterExpression);
        }

        /// <summary>
        /// 生成类似于p=>p.values.Contains("xxx");的lambda表达式
        /// parameterExpression标识p，propertyName表示values，propertyValue表示"xxx",methodName表示Contains
        /// 仅处理p的属性类型为string这种情况
        /// </summary>
        /// <param name="methodName"></param>
        /// <param name="propertyName"></param>
        /// <param name="propertyValue"></param>
        /// <param name="parameterExpression"></param>
        /// <returns></returns>
        private static MethodCallExpression GetMethodExpression(string methodName, string propertyName, string propertyValue, ParameterExpression parameterExpression)
        {
            var propertyExpression = Expression.Property(parameterExpression, propertyName);
            MethodInfo method = typeof(string).GetMethod(methodName, new[] { typeof(string) });
            var someValue = Expression.Constant(propertyValue, typeof(string));
            return Expression.Call(propertyExpression, method, someValue);
        }
    }

    public static class LinqBuilder
    {
        /// <summary>
        /// 默认True条件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Expression<Func<T, bool>> True<T>()
        {
            return f => true;
        }

        /// <summary>
        /// 默认False条件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Expression<Func<T, bool>> False<T>()
        {
            return f => false;
        }

        public static Expression<T> Compose<T>(this Expression<T> first, Expression<T> second, Func<Expression, Expression, Expression> merge)
        {
            var map = first.Parameters.Select((f, i) => new { f, s = second.Parameters[i] }).ToDictionary(p => p.s, p => p.f);
            var secondBody = ParameterRebinder.ReplaceParameters(map, second.Body);
            return Expression.Lambda<T>(merge(first.Body, secondBody), first.Parameters);
        }

        /// <summary>
        /// 拼接 OR 条件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="exp"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        public static Expression<Func<T, bool>> Or<T>(this Expression<Func<T, bool>> exp, Expression<Func<T, bool>> condition)
        {
            //var inv = Expression.Invoke(condition, exp.Parameters.Cast<Expression>());
            //return Expression.Lambda<Func<T, bool>>(Expression.Or(exp.Body, inv.Expression), exp.Parameters);
            return exp.Compose(condition, Expression.Or);
        }

        /// <summary>
        /// 拼接And条件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="exp"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        public static Expression<Func<T, bool>> And<T>(this Expression<Func<T, bool>> exp, Expression<Func<T, bool>> condition)
        {
            //var inv = Expression.Invoke(condition, exp.Parameters.Cast<Expression>());
            //return Expression.Lambda<Func<T, bool>>(Expression.And(exp.Body, inv), exp.Parameters);
            return exp.Compose(condition, Expression.And);
        }
    }

    public class ParameterRebinder : ExpressionVisitor
    {
        private readonly Dictionary<ParameterExpression, ParameterExpression> map;

        public ParameterRebinder(Dictionary<ParameterExpression, ParameterExpression> map)
        {
            this.map = map ?? new Dictionary<ParameterExpression, ParameterExpression>();
        }

        public static Expression ReplaceParameters(Dictionary<ParameterExpression, ParameterExpression> map, Expression exp)
        {
            return new ParameterRebinder(map).Visit(exp);
        }

        protected override Expression VisitParameter(ParameterExpression p)
        {
            ParameterExpression replacement;
            if (map.TryGetValue(p, out replacement))
            {
                p = replacement;
            }
            return base.VisitParameter(p);
        }
    }
}