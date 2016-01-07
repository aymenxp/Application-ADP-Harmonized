using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HarmonizedApp.Manager.BO;
using System.Data.OleDb;


namespace HarmonizedApp.Manager.DAL
{
   public class clsRubricStructureDAL:clsBD
    {

     public void Add(clsRubricStructureBO Rub)
       {
            OleDbConnection con = null;
           
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();

                OleDbCommand cmd = new OleDbCommand("insert into RubriqueStructureGrossToNet (Id,IdSoc,Label,NomColonne,ColonneType,StartString,EndString,ConvertOut,ReplaceString,Order)"
                    + "values (@Id,@IdSoc,@Label,@NomColonne,@ColonneType,@StartString,@EndString,@ConvertOut,@ReplaceString,@Order)"
                    , con);

                cmd.Parameters.Add(new OleDbParameter("@Id", Rub.Id));
                cmd.Parameters.Add(new OleDbParameter("@IdSoc", Rub.IdSoc));
                cmd.Parameters.Add(new OleDbParameter("@Label", Rub.Label));
                cmd.Parameters.Add(new OleDbParameter("@NomColonne", Rub.Label));
                cmd.Parameters.Add(new OleDbParameter("@StartString", Rub.StartString));
                cmd.Parameters.Add(new OleDbParameter("@EndString", Rub.EndString));
                cmd.Parameters.Add(new OleDbParameter("@ColonneType", Rub.ColonneType));
                cmd.Parameters.Add(new OleDbParameter("@ConvertOut", Rub.ConvertOut));
                cmd.Parameters.Add(new OleDbParameter("@ReplaceString", Rub.ReplaceString));
                cmd.Parameters.Add(new OleDbParameter("@Order", Rub.ReplaceString));



                cmd.ExecuteNonQuery();
             
            }
            catch
            {

            }
            finally
            {
                con.Close();
                con.Dispose();
              
            }


       }


      

       public void Update(clsRubricStructureBO Rub)
       {
           OleDbConnection con = null;
         
           try
           {
               con = new OleDbConnection(connection_string);
               con.Open();

               OleDbCommand cmd = new OleDbCommand("update  RubriqueStructureGrossToNet "
                   + " set IdSoc=@IdSoc,Label=@Label,NomColonne=@NomColonne,ColonneType=@ColonneType,StartString=@StartString,EndString=@EndString,ConvertOut=@ConvertOut,ReplaceString=@ReplaceString)"
                   + " where Id=@Id"
                   , con);

               cmd.Parameters.Add(new OleDbParameter("@Id", Rub.Id));
               cmd.Parameters.Add(new OleDbParameter("@IdSoc", Rub.IdSoc));
               cmd.Parameters.Add(new OleDbParameter("@Label", Rub.Label));
               cmd.Parameters.Add(new OleDbParameter("@NomColonne", Rub.Label));
               cmd.Parameters.Add(new OleDbParameter("@StartString", Rub.StartString));
               cmd.Parameters.Add(new OleDbParameter("@EndString", Rub.EndString));
               cmd.Parameters.Add(new OleDbParameter("@ColonneType", Rub.ColonneType));
               cmd.Parameters.Add(new OleDbParameter("@ConvertOut", Rub.ConvertOut));
               cmd.Parameters.Add(new OleDbParameter("@ReplaceString", Rub.ReplaceString));
               cmd.Parameters.Add(new OleDbParameter("@Order", Rub.ReplaceString));



               cmd.ExecuteNonQuery();
             

           }
           catch
           {

           }
           finally
           {
               
               con.Close();
               con.Dispose();
               
           }

       }

       public void Delete(clsRubricStructureBO Rub)
       {
           OleDbConnection con = null;

           try
           {
               con = new OleDbConnection(connection_string);
               con.Open();

               OleDbCommand cmd = new OleDbCommand("delete from RubriqueStructureGrossToNet "
                   + " where Id=@Id"
                   , con);

               cmd.Parameters.Add(new OleDbParameter("@Id", Rub.Id));
             
               cmd.ExecuteNonQuery();


           }
           catch
           {

           }
           finally
           {

               con.Close();
               con.Dispose();

           }
       }

       public clsRubricStructureBO GetDetails(int Id)
       {
           OleDbConnection con = null;
           clsRubricStructureBO rub = new clsRubricStructureBO();

           try
           {
               con = new OleDbConnection(connection_string);
               con.Open();

               OleDbCommand cmd = new OleDbCommand("select Id,IdSoc,Label,NomColonne,ColonneType,StartString,EndString,ConvertOut,ReplaceString,Order from RubriqueStructureGrossToNet "
                   , con);

              OleDbDataReader dr = cmd.ExecuteReader();

              if (dr.Read())
              {
                  rub.Id = int.Parse(dr["Id"].ToString());
                  rub.IdSoc = int.Parse(dr["IdSoc"].ToString());
                  rub.Label = dr["Label"].ToString();
                  rub.NomColonne = dr["NomColonne"].ToString();
                  rub.ReplaceString = dr["ReplaceString"].ToString();
                  rub.StartString = int.Parse(dr["StartString"].ToString());
                  rub.EndString = int.Parse(dr["EndString"].ToString());
                  rub.ColonneType = dr["ColonneType"].ToString();
                  rub.Order = int.Parse(dr["Order"].ToString());
              }

           }
           catch
           {

           }
           finally
           {
               con.Close();
               con.Dispose();

           }



           return rub;
       }


       public List<clsRubricStructureBO> GetListRubricsByIdSoc(int IdSoc)
       {
           OleDbConnection con = null;
           clsRubricStructureBO rub;
           List<clsRubricStructureBO> List = new List<clsRubricStructureBO>();

           try
           {
               con = new OleDbConnection(connection_string);
               con.Open();

               OleDbCommand cmd = new OleDbCommand("select Id,IdSoc,Label,NomColonne,ColonneType,StartString,EndString,ConvertOut,ReplaceString,Order from RubriqueStructureGrossToNet "
                   + " where IdSoc = " + IdSoc
                   + " Order by Oder "
                   , con);

               OleDbDataReader dr = cmd.ExecuteReader();

               while (dr.Read())
               {
                   rub = new clsRubricStructureBO();

                   rub.Id = int.Parse(dr["Id"].ToString());
                   rub.IdSoc = int.Parse(dr["IdSoc"].ToString());
                   rub.Label = dr["Label"].ToString();
                   rub.NomColonne = dr["NomColonne"].ToString();
                   rub.ReplaceString = dr["ReplaceString"].ToString();
                   rub.StartString = int.Parse(dr["StartString"].ToString());
                   rub.EndString = int.Parse(dr["EndString"].ToString());
                   rub.ColonneType = dr["ColonneType"].ToString();
                   rub.Order = int.Parse(dr["Order"].ToString());

                   List.Add(rub);
               }


           }
           catch
           {

           }
           finally
           {
               con.Close();
               con.Dispose();

           }



           return List;
       }




    }
}
