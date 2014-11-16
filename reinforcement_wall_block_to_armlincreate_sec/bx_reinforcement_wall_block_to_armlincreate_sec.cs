using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using App = Autodesk.AutoCAD.ApplicationServices;
using cad = Autodesk.AutoCAD.ApplicationServices.Application;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Gem = Autodesk.AutoCAD.Geometry;
using Rtm = Autodesk.AutoCAD.Runtime;

// [assembly: Rtm.CommandClass(typeof(MyClassSerializer.Commands))]

namespace reinforcement_wall_block_to_armlincreate_sec
{
    public class Commands
    {
        //Проверка блоков отверстий, в Autocad, на попадание в отметку этажа
        [Rtm.CommandMethod("bx_reinforcement_wall_block_to_armlincreate_sec")]
        static public void bx_reinforcement_wall_block_to_armlincreate_sec()
        {
            // Получение текущего документа и базы данных
            App.Document acDoc = App.Application.DocumentManager.MdiActiveDocument;
            Db.Database acCurDb = acDoc.Database;
            Ed.Editor acEd = acDoc.Editor;

            // старт транзакции
            using (Db.Transaction acTrans = acCurDb.TransactionManager.StartOpenCloseTransaction())
            {

                Db.TypedValue[] acTypValAr = new Db.TypedValue[1];
                acTypValAr.SetValue(new Db.TypedValue((int)Db.DxfCode.Start, "INSERT"), 0);
                Ed.SelectionFilter acSelFtr = new Ed.SelectionFilter(acTypValAr);


                Ed.PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection(acSelFtr);
                if (acSSPrompt.Status == Ed.PromptStatus.OK)
                {

                    // Открытие таблицы Блоков для чтения
                    Db.BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, Db.OpenMode.ForRead) as Db.BlockTable;

                    // Открытие записи таблицы Блоков пространства Модели для записи
                    Db.BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[Db.BlockTableRecord.ModelSpace],
                                                                                    Db.OpenMode.ForWrite) as Db.BlockTableRecord;



                    Ed.SelectionSet acSSet = acSSPrompt.Value;
                    foreach (Ed.SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            Db.Entity acEnt = acTrans.GetObject(acSSObj.ObjectId,
                                                     Db.OpenMode.ForRead) as Db.Entity;
                            if (acEnt != null)
                            {
                                if (acEnt is Db.BlockReference)
                                {
                                    Db.BlockReference acBlRef = (Db.BlockReference)acEnt;


                                    // тут нужна проверка имени блока.
                                    //тут выясняю истинное имя блока для последующего обновления атрибутов.
                                    //Проверяю является ли выделенный блок динамическим
                                    //Получаю настоящие/родное имя динамического блока

                                    Db.BlockTableRecord blr = (Db.BlockTableRecord)acTrans.GetObject(acBlRef.DynamicBlockTableRecord,
                                                                                                    Db.OpenMode.ForRead);

                                        Db.BlockTableRecord blr_nam = (Db.BlockTableRecord)acTrans.GetObject(blr.ObjectId,
                                                                                                    Db.OpenMode.ForRead);
                                        // тут лежит имя блока, в том числе динамческого блока
                                        String acBlock_nam = blr_nam.Name.ToUpper();


                                        //Теперь вот этот вот фрагмент кода на VB.NEt  надо переписать на С#
                                        /*
                               If acBlock_nam.ToUpper Like "Отв с мар*".ToUpper Or
                                    acBlock_nam.ToUpper = "Otverstie".ToUpper Then
                                    ...
                                End If
                                         */

                                        if (acBlock_nam.Trim().Contains("reinforcement_wall_plus".ToUpper()))                                            
                                        {
                                            int KR_number = 0;
                                            int Raz_number = 0;

                                            int KR_Step = 0;
                                            int KR_Poz1_diam = 0;

                                            int PROD_poz = 0;
                                            int PROD_diam = 0;

                                            Double B_st = 0;
                                            Double H_st = 0;

                                            Double OTM_Niza = 0.000;

                                            Double ZS_niz = 0;
                                            Double ZS_verh = 0;
                                            Double ZS_lev = 0;
                                            Double ZS_prav = 0;

                                            Double Poper_step = 0;
                                            Double Distance1 = 0;



                                            // Тут еще нужно считать динамический параметр "Ширина"
                                            //и проверить отверстие на квадратность

                                            Db.DynamicBlockReferencePropertyCollection acBlockDynProp =
                                                acBlRef.DynamicBlockReferencePropertyCollection;
                                            if (acBlockDynProp != null)
                                            {
                                                foreach (Db.DynamicBlockReferenceProperty obj in acBlockDynProp)
                                                {
                                                    //		obj.PropertyName	"KR_number"	string

                                                     switch (obj.PropertyName)
                                                     {
                                                         case "KR_number":
                                                             {
                                                                 KR_number = (int)Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                 break;
                                                             }
                                                         case "Raz_number":
                                                             {
                                                                 Raz_number = (int)Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                 break;
                                                             }
                                                         case "KR_step":
                                                             {
                                                                 KR_Step = (int)Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                 break;
                                                             }

                                                         case "KR_Poz1_diam":
                                                             {
                                                                 KR_Poz1_diam = (int)Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                 break;
                                                             }

                                                         case "PROD_poz":
                                                             {
                                                                 PROD_poz = (int)Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                 break;
                                                             }

                                                         case "PROD_diam":
                                                             {
                                                                 PROD_diam = (int)Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                 break;
                                                             }

                                                        case "B_st":
                                                            {
                                                                B_st = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }

                                          
                                                        case "H_St":
                                                            {
                                                                H_st = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }
                                                    
                                                        case "OTM_Niza":
                                                            {
                                                                OTM_Niza = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }

                                                        case "ZS_niz":
                                                            {
                                                                ZS_niz = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }

                                                        case "ZS_verh":
                                                            {
                                                                ZS_verh = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }

                                                        case "ZS_lev":
                                                            {
                                                                ZS_lev = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }
                                                        case "ZS_prav":
                                                            {
                                                                ZS_prav = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }

                                                        //Poper_step
                                                        case "Poper_step":
                                                            {
                                                                Poper_step = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }

                                                        //Distance1
                                                        case "Distance1":
                                                            {
                                                                Distance1 = Math.Round(Double.Parse(obj.Value.ToString()), 0);
                                                                break;
                                                            }
                                                         
                                                        default:
                                                            break;
                                                     }

                                                }
                                            }

// Вот тут выводим сформированное описание



                                            //string raz = (Raz_number + "-" + Raz_number + "\n" +
                                            //    "| " + PROD_diam + " " + (int)(Poper_step/200+1)*2 + " pm\n" + 
                                            //    "_ Кр " + KR_number + " " + KR_Step + "\n" +
                                            //    "S " + (B_st*H_st)/(1000*1000) + " M B25W8").Trim()  ;                                              ;

                                            //string kr = ("Кр " + KR_number + "\n" +
                                            //    "| " + KR_Poz1_diam + " " + (int)(Distance1 - ZS_niz + ZS_verh) + " 2\n" +
                                            //    "_ 6 " + (B_st - 20*2) + " " + (int)(Poper_step / 200 + 1)).Trim();

                                            string p = "\\P";

                                            string raz = (Raz_number + "-" + Raz_number + p +
                                            "| " + PROD_diam + " " + (int)(Poper_step / 200 + 1) * 2 + " pm" + p +
                                            "_ Кр " + KR_number + " " + KR_Step + p +
                                            "S " + (B_st * H_st) / (1000 * 1000) + " M B25W8").Trim(); ;

                                            string kr = ("Кр " + KR_number + p +
                                                "| " + KR_Poz1_diam + " " + (int)(Distance1 - ZS_niz + ZS_verh) + " 2" + p +
                                                "_ 6 " + (B_st - 20 * 2) + " " + (int)(Poper_step / 200 + 1)).Trim();


                                            Ed.PromptPointOptions pointOpt = new Ed.PromptPointOptions("\n Точка вставки описания:");
                                            Ed.PromptPointResult pointRes = acEd.GetPoint(pointOpt);
                                            if (pointRes.Status != Ed.PromptStatus.OK)
                                            {
                                                return;
                                            }

                                            Db.MText mtxt1 = new Db.MText();
                                            mtxt1.Contents = raz;
                                            mtxt1.TextHeight = 50;
                                            mtxt1.ColorIndex = Raz_number;

                                            mtxt1.Location = pointRes.Value;
                                            mtxt1.SetDatabaseDefaults();
                                            acBlkTblRec.AppendEntity(mtxt1);
                                            acTrans.AddNewlyCreatedDBObject(mtxt1, true);

                                            Db.MText mtxt2 = new Db.MText();
                                            mtxt2.Contents = kr;
                                            mtxt2.TextHeight = 50;
                                            mtxt2.ColorIndex = KR_number;
                                            mtxt2.Location = new Gem.Point3d(pointRes.Value.X, pointRes.Value.Y - 400, 0);
                                            mtxt2.SetDatabaseDefaults();
                                            acBlkTblRec.AppendEntity(mtxt2);
                                            acTrans.AddNewlyCreatedDBObject(mtxt2, true);

                                            //// Тут вывести количество обработтаных блоков  и кол квадратных блоков
                                            acEd.WriteMessage("\n " + raz);
                                            acEd.WriteMessage("\n " + kr);


                                        }  //Проверка имени блока



                                }   //Проверка, что объект это ссылка на блок
                                //acEnt.ColorIndex = 3;
                            }
                        }
                    }
                }

                acTrans.Commit();


            }


            ////// Тут вывести количество обработтаных блоков  и кол квадратных блоков
            //acEd.WriteMessage("\n Количество обработтаных блоков: " + ObjSelCount.ToString());
            //acEd.WriteMessage("\n Количество КВАДРАТНЫХ блоков: " + squareCount.ToString());


        }
    }
}
