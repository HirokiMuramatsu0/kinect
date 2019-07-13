using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;
using Microsoft.Kinect.VisualGestureBuilder;
using Excel = Microsoft.Office.Interop.Excel;


namespace sotukenn
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        //kinect
        KinectSensor kinect;
        //kinect 骨格検知
        BodyFrameReader bodyFrameReader;
        Body[] bodies;
        //kinect ジェスチャー検知
       // VisualGestureBuilderDatabase gestureDatabase;
       // VisualGestureBuilderFrameReader[] gestureFrameReaders;
       // IReadOnlyList<Gesture> gestures;

        int BODY_COUNT;
        int NUMBER_OF_FRAME = 80;

        //for Position Check
        double z;
        double z2;
        double y;
        double y2;
        int m = 0;
        int s = 0;
        int a = 0;

        CheckData[] checkDatas;

        //Stopwatch
        System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
        private static DateTime UNIX_EPOCH = new DateTime(1970, 1, 1, 0, 0, 0, 0);

        //writeDate
        private Excel.Application mExcel;
        private Excel.Workbook mWorkbook;
        private Excel.Worksheet mSheet;

        //test
        private bool fallDown = false;

        ColorFrameReader colorFrameReader;
        FrameDescription colorFrameDesc;

        ColorImageFormat colorFormat = ColorImageFormat.Bgra;

        // WPF
        WriteableBitmap colorBitmap;
        byte[] colorBuffer;
        int colorStride;
        Int32Rect colorRect;


        ///////////////////////////////////////////Main Functions //////////////////////////////////////////////////////////////////

        public MainWindow()
        {
            InitializeComponent();
            
        }

       

        private void Window_Loaded(object sensor, RoutedEventArgs e)
        {


            try
            {
                kinect = KinectSensor.GetDefault();
                if (kinect == null)
                {
                    throw new Exception("Kinectを開けません");
                }

                kinect.Open();

                BODY_COUNT = kinect.BodyFrameSource.BodyCount;

                //ボディーリーダーを開く
                bodyFrameReader = kinect.BodyFrameSource.OpenReader();
                bodyFrameReader.FrameArrived += bodyFrameReader_FrameArrived;


                //+++++++++++++++++++++++++++++++
                // カラー画像の情報を作成する(BGRAフォーマット)
                colorFrameDesc = kinect.ColorFrameSource.CreateFrameDescription(colorFormat);

                //// カラーリーダーを開く
                colorFrameReader = kinect.ColorFrameSource.OpenReader();
                colorFrameReader.FrameArrived += colorFrameReader_FrameArrived;

                //// カラー用のビットマップを作成する
                colorBitmap = new WriteableBitmap(
                                    colorFrameDesc.Width, colorFrameDesc.Height,
                                    96, 96, PixelFormats.Bgra32, null);
                colorStride = colorFrameDesc.Width * (int)colorFrameDesc.BytesPerPixel;
                colorRect = new Int32Rect(0, 0,
                                    colorFrameDesc.Width, colorFrameDesc.Height);
                colorBuffer = new byte[colorStride * colorFrameDesc.Height];
                ImageColor.Source = colorBitmap;
                //+++++++++++++++++++++++++++++++++++++++++++++


                //Bodyを入れる配列を作る
                bodies = new Body[BODY_COUNT];
                checkDatas = new CheckData[BODY_COUNT];





            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Close();
            }
        }

       private void Window_Closing(object sensor,
               System.ComponentModel.CancelEventArgs e)
        {
            if (bodyFrameReader != null)
            {
                bodyFrameReader.Dispose();
                bodyFrameReader = null;

                if (colorFrameReader != null)
                {
                    colorFrameReader.Dispose();
                    colorFrameReader = null;
                }
            }
            if (kinect != null)
            {
                kinect.Close();
                kinect = null;
            }

            //closeExcel();
        }

        ///////////////////////////////////////////Functions //////////////////////////////////////////////////////////////////

        void bodyFrameReader_FrameArrived(object sender, BodyFrameArrivedEventArgs e)
        {
            //Console.WriteLine("bodyFrameReader_FrameArrived");
            //ここがループし続ける
            UpdateBodyFrame(e);
           
            DrawBodyFrame();
        }
        //ボディの更新
        private void UpdateBodyFrame(BodyFrameArrivedEventArgs e)
        {
            using (var bodyFrame = e.FrameReference.AcquireFrame())
            {
                if (bodyFrame == null)
                {
                    return;
                }
                //ボディデータを取得する
                bodyFrame.GetAndRefreshBodyData(bodies);
            }
        }
           
        //+++++++++++++++++++++++++++++++
        void colorFrameReader_FrameArrived(object sender, ColorFrameArrivedEventArgs e)
        {
            UpdateColorFrame(e);
            DrawColorFrame();
        }

        private void UpdateColorFrame(ColorFrameArrivedEventArgs e)
        {
            // カラーフレームを取得する
            using (var colorFrame = e.FrameReference.AcquireFrame())
            {
                if (colorFrame == null)
                {
                    return;
                }

                // BGRAデータを取得する
                colorFrame.CopyConvertedFrameDataToArray(
                                            colorBuffer, colorFormat);
            }
        }

        private void DrawColorFrame()
        {
            // ビットマップにする
            colorBitmap.WritePixels(colorRect, colorBuffer,
                                            colorStride, 0);
        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        //ボディの表示
        private void DrawBodyFrame()
        {
            CanvasBody.Children.Clear();
            if (bodies.Length == 0)
                return;

            //Body body = bodies[0];

            //追跡しているBodyのみループする
            foreach (var body in bodies.Where(b => b.IsTracked))
            {
                foreach (var joint in body.Joints)
                {

                    //関節の位置が追跡状態
                    if (joint.Value.TrackingState == TrackingState.Tracked)
                    {
                        DrawEllipse(joint.Value, 10, Brushes.Blue);
                    }
                    //関節の位置が推測状態
                    else if (joint.Value.TrackingState == TrackingState.Inferred)
                    {
                        DrawEllipse(joint.Value, 10, Brushes.Yellow);
                    }
                }
            }
        }

        private void DrawEllipse(Joint joint, int R, Brush brush)
        {
            var ellipse = new Ellipse()
            {
                Width = R,
                Height = R,
                Fill = brush,
            };

            //カメラ座標をdepth座標に変換する
            var point = kinect.CoordinateMapper.MapCameraPointToDepthSpace(
                                                                    joint.Position);
            if ((point.X < 0) || (point.Y < 0))
            {
                return;
            }

            //Depth座標系で円を配置する
            Canvas.SetLeft(ellipse, point.X - (R / 2));
            Canvas.SetTop(ellipse, point.Y - (R / 2));

            CanvasBody.Children.Add(ellipse);
        }

       

        private void CheckJointPositions4()
        {
            if (bodies.Length == 0)
                return;

            Body body = bodies[0];

            double headz = Math.Round(body.Joints[JointType.Head].Position.Z * 30, 0);
            //      double handRight = Math.Round(body.Joints[JointType.HandRight].Position.Z * 30, 0);
            //      double handLeft = Math.Round(body.Joints[JointType.HandLeft].Position.Z * 30, 0);
            double spinebasez = Math.Round(body.Joints[JointType.SpineBase].Position.Z * 30, 0);
            //      double kneeLeft = Math.Round(body.Joints[JointType.KneeLeft].Position.Z * 30, 0);
            //      double kneeRight = Math.Round(body.Joints[JointType.KneeRight].Position.Z * 30, 0);
            //      double footLeft = Math.Round(body.Joints[JointType.FootLeft].Position.Z * 30, 0);
            //      double footRight = Math.Round(body.Joints[JointType.FootRight].Position.Z * 30, 0);
            double heady = Math.Round(body.Joints[JointType.Head].Position.Y * 30, 0);
            double spinebasey = Math.Round(body.Joints[JointType.SpineBase].Position.Y * 30, 0);
            double z1 = Math.Abs(headz - spinebasez);
            double y1 = Math.Abs(heady - spinebasey);

            z = Math.Abs(z1 - z2);
            z2 = z1;
            y = Math.Abs(y1 - y2);
            y2 = y1;


            //     Console.WriteLine("head = {0}, handRight = {1}, handLeft = {2}, kneeRight = {3}, kneeLeft = {4}, spinebase = {5}, footLeft = {6}, footRight = {7}", head, handRight, handLeft, kneeRight, kneeLeft, spinebase, footLeft, footRight);
            //   Console.WriteLine("z1 = {0}", z1);
            //Console.WriteLine("z = {0}", z);
            if (z != 0)
            {
                m = m + 1;
                // Console.WriteLine("m = {0}", m);
                if (m > 3)//しきい値決める際に人間の転ぶ速度の平均を計算（Kinectは30fpsで値を読む）
                {
                    if (z1 > 10)//この条件はいい感じ（起き上がるときと差別化）
                    {
                        Title2.Content = "異常を検知しました";
                        //     a = a + 1;
                        //     Console.WriteLine("a = {0}", a);


                    }
                    /*    else
                        {
                            Title2.Content = "異常なし";

                        }*/
                }
                /*       else if(j > 5)//かがむ姿勢の判別どーする？

                       {

                       }*/

            }
            else
            {
                s = s + 1;
                //Console.WriteLine("s = {0}", s);
                if (s > 1)
                {


                    m = 0;
                    s = 0;
                    return;

                }




            }


            //      Title1.Content = "異常なし";
            //    Title2.Content = "異常なし";
            //   return;





        }

        private void reset_Click(object sender, RoutedEventArgs e)
        {
            //Title1.Content = "異常なし";
            //Title2.Content = "異常なし";
            NotificationBlock.Content = "転倒を検知しました。";
            NotificationBlock.Background = Brushes.Red;
        }


        private void check_BodySpead(int index)
        {
            Console.WriteLine("CheckHeadSpead");

            CheckData checkData = checkDatas[index];
            double[] headXInfo = checkData.headXInfo;
            double[] headYInfo = checkData.headYInfo;
            double[] headZInfo = checkData.headZInfo;
            double[] middleXInfo = checkData.middleXInfo;
            double[] middleYInfo = checkData.middleYInfo;
            double[] middleZInfo = checkData.middleZInfo;
            long[] time = checkData.time;



            Body body = bodies[index];
            if (!body.IsTracked)
            {
                return;
            }

            double headX = Math.Round(body.Joints[JointType.Head].Position.X * 30, 0);
            double headY = Math.Round(body.Joints[JointType.Head].Position.Y * 30, 0);
            double headZ = Math.Round(body.Joints[JointType.Head].Position.Z * 30, 0);

            double middleX = Math.Round(body.Joints[JointType.SpineMid].Position.X * 30, 0);
            double middleY = Math.Round(body.Joints[JointType.SpineMid].Position.Y * 30, 0);
            double middleZ = Math.Round(body.Joints[JointType.SpineMid].Position.Z * 30, 0);

            double oldHeadX = headXInfo.First();
            double oldHeadY = headYInfo.First();
            double oldHeadZ = headZInfo.First();

            double oldMiddleX = middleXInfo.First();
            double oldMiddleY = middleYInfo.First();
            double oldMiddleZ = middleZInfo.First();

            double deltaHeadX = System.Math.Pow((headX - oldHeadX), 2);
            double deltaHeadY = System.Math.Pow((headY - oldHeadY), 2);
            double deltaHeadZ = System.Math.Pow((headZ - oldHeadZ), 2);

            double deltaMiddleX = System.Math.Pow((middleX - oldMiddleX), 2);
            double deltaMiddleY = System.Math.Pow((middleY - oldMiddleY), 2);
            double deltaMiddleZ = System.Math.Pow((middleZ - oldMiddleZ), 2);


            double headMove3D = System.Math.Sqrt(deltaHeadX + deltaHeadY + deltaHeadZ);
            double middleMove3D = System.Math.Sqrt(deltaMiddleX + deltaMiddleY + deltaMiddleZ);

            double difTime = GetUnixTime(DateTime.Now) - time.First();

            double speedHead = headMove3D / difTime;
            double speedMiddle = middleMove3D / difTime;


            Console.WriteLine("now headX:{0}", headX);
            Console.WriteLine("old headX:{0}", oldHeadX);
            Console.WriteLine("now headY:{0}", headY);
            Console.WriteLine("old headY:{0}", oldHeadY);
            Console.WriteLine("now headZ:{0}", headZ);
            Console.WriteLine("old headZ:{0}", oldHeadZ);



            Console.WriteLine("******Move Head 3D : {0}******", headMove3D);
            Console.WriteLine("******Move Middle 3D : {0}******", middleMove3D);
            Console.WriteLine("******Dif Time : {0}******", difTime);
            Console.WriteLine("******Speed Head : {0}******", speedHead);
            Console.WriteLine("******Speed Middle : {0}******", speedMiddle);


            double[] writeData = { headMove3D, middleMove3D, difTime, speedHead, speedMiddle };



            if (headZ < 120 && headZ > 10)//ボーントラッキングの範囲外に出たときにheadｚがおかしくなって誤検知が起きるのでここで調整
            {
                if (headMove3D > 38 && middleMove3D > 36)
                {
                    Console.WriteLine("************Move Fall Down**************");

                }
                if (speedHead > 0.016 && speedMiddle > 0.015)
                {
                    Console.WriteLine("************Speed Fall Down**************");
                    this.fallDown = true;
                    NotificationBlock.Background = Brushes.Red;

                }
            }

            //writeDataToExcel(writeData);

            //this.Close();


        }

        //bodyDataの入れ替え
        private void updateCheckData(int index)
        {
            Console.WriteLine("updateCheckData");

            CheckData checkData = checkDatas[index];
            int numberForArraySort = NUMBER_OF_FRAME - 1;

            //time
            long dt = GetUnixTime(DateTime.Now);

            //このindexのnowでーた取得
            Body body = bodies[index];
            double nowHeadX = Math.Round(body.Joints[JointType.Head].Position.X * 30, 0);
            double nowHeadY = Math.Round(body.Joints[JointType.Head].Position.Y * 30, 0);
            double nowHeadZ = Math.Round(body.Joints[JointType.Head].Position.Z * 30, 0);

            double nowMiddleX = Math.Round(body.Joints[JointType.SpineMid].Position.X * 30, 0);
            double nowMiddleY = Math.Round(body.Joints[JointType.SpineMid].Position.Y * 30, 0);
            double nowMiddleZ = Math.Round(body.Joints[JointType.SpineMid].Position.Z * 30, 0);



            //初期の時
            if (checkData == null)
            {
                checkData = new CheckData();

                if (index == 0)
                {
                    sw.Start();
                    Console.WriteLine("***?????? ストップウォッチ開始");
                }

            }

            //data代入
            double[] headXInfo = checkData.headXInfo;
            double[] headYInfo = checkData.headYInfo;
            double[] headZInfo = checkData.headZInfo;

            double[] middleXInfo = checkData.middleXInfo;
            double[] middleYInfo = checkData.middleYInfo;
            double[] middleZInfo = checkData.middleZInfo;


            long[] time = checkData.time;

            int length = headZInfo.Length;

            Console.WriteLine("***length:{0}", length);


            if (length < NUMBER_OF_FRAME)//配列埋まりきっていない時の処理-配列の長さを図ってその次のところに追加
            {
                Array.Resize(ref headXInfo, length + 1);
                Array.Resize(ref headYInfo, length + 1);
                Array.Resize(ref headZInfo, length + 1);
                Array.Resize(ref middleXInfo, length + 1);
                Array.Resize(ref middleYInfo, length + 1);
                Array.Resize(ref middleZInfo, length + 1);

                Array.Resize(ref time, length + 1);
                //nowDataを末尾に追加
                headXInfo[length] = nowHeadX;
                headYInfo[length] = nowHeadY;
                headZInfo[length] = nowHeadZ;

                middleXInfo[length] = nowMiddleX;
                middleYInfo[length] = nowMiddleY;
                middleZInfo[length] = nowMiddleZ;


                time[length] = dt;
            }
            else//配列が埋まりきっているとき-配列の中身入れ替え
            {
                //Info中身入れ替え
                for (int i = 0; i < numberForArraySort; ++i)
                {
                    headXInfo[i] = headXInfo[i + 1];
                    headYInfo[i] = headYInfo[i + 1];
                    headZInfo[i] = headZInfo[i + 1];
                    middleXInfo[i] = middleXInfo[i + 1];
                    middleYInfo[i] = middleYInfo[i + 1];
                    middleZInfo[i] = middleZInfo[i + 1];
                    time[i] = time[i + 1];
                }
                headXInfo[numberForArraySort] = nowHeadX;//nowDataをいれる
                headYInfo[numberForArraySort] = nowHeadY;
                headZInfo[numberForArraySort] = nowHeadZ;
                middleXInfo[numberForArraySort] = nowMiddleX;//nowDataをいれる
                middleYInfo[numberForArraySort] = nowMiddleY;
                middleZInfo[numberForArraySort] = nowMiddleZ;
                time[numberForArraySort] = dt;

                if (index == 0)
                {
                    sw.Stop();
                    Console.WriteLine("***?????? 時間:{0}", sw.Elapsed);
                }
            }


            foreach (int i in headZInfo)
            {
                // Console.WriteLine(i);
            }

            //データ保存
            checkData.middleXInfo = middleXInfo;
            checkData.middleYInfo = middleYInfo;
            checkData.middleZInfo = middleZInfo;

            checkData.headXInfo = headXInfo;
            checkData.headYInfo = headYInfo;
            checkData.headZInfo = headZInfo;
            checkData.time = time;

            checkDatas[index] = checkData;

        }


        void printBodyData(int index)
        {
            Body body = bodies[index];
            if (!body.IsTracked)
            {
                return;
            }

            double headX = body.Joints[JointType.Head].Position.X;
            Console.WriteLine("頭の位置{0}", headX);
        }


        ///////////////////////////////////////////Blue Touce//////////////////////////////////////////////////////////////////

        /*      private void cmdConnect_Click(object sender, RoutedEventArgs e)
              {
                  try
                  {
                      if (SPBluetooth.GetInsntace().IsOpen)
                      {
                          SPBluetooth.GetInsntace().Close();
                          cmdConnect.Content = "Connect";
                      }
                      else
                      {
                          SPBluetooth.GetInsntace().PortName = comboCOMMPorts.SelectedItem.ToString().Trim();

                          // configs
                          SPBluetooth.GetInsntace().BaudRate = 115200;
                          SPBluetooth.GetInsntace().DataBits = 8;

                          SPBluetooth.GetInsntace().StopBits = System.IO.Ports.StopBits.One;
                          SPBluetooth.GetInsntace().Parity = System.IO.Ports.Parity.None;

                          SPBluetooth.GetInsntace().Open();
                          cmdConnect.Content = "Disconnect";
                      }
                  }
                  catch (Exception)
                  {
                      MessageBox.Show("Comm port Not Available \n コンポート接続されていません", "Comm port Error", MessageBoxButton.OK, MessageBoxImage.Error);
                  }
              }

              private void cmdEchoOn_Click(object sender, RoutedEventArgs e)
              {
                  SPBluetooth.GetInsntace().WriteLine("+");
              }

              private void cmdEchoOff_Click(object sender, RoutedEventArgs e)
              {
                  SPBluetooth.GetInsntace().WriteLine("+");
              }

              private void cmdSearch_Click(object sender, RoutedEventArgs e)
              {
                  SPBluetooth.GetInsntace().WriteLine("F");
              }

              private void cmdStopSearch_Click(object sender, RoutedEventArgs e)
              {
                  SPBluetooth.GetInsntace().WriteLine("X");
              }

              private void cmdsconnectBlootooth_Click(object sender, RoutedEventArgs e)
              {
                  SPBluetooth.GetInsntace().WriteLine("E,0,001EC0462DE8");
              }*/

        private static long GetUnixTime(DateTime targetTime)
        {
            targetTime = targetTime.ToUniversalTime();
            TimeSpan elapsedTime = targetTime - UNIX_EPOCH;

            return (long)elapsedTime.TotalMilliseconds;
        }
        private void excelSetup()
        {
            this.mExcel = new Excel.Application();
            this.mExcel.Visible = false;

            this.mWorkbook = (Excel.Workbook)(mExcel.Workbooks.Open(
                @"C:\Users\ag15040\Desktop\FallDownData.xlsx",
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing));

            // シートを取得する
            this.mSheet = (Excel.Worksheet)this.mWorkbook.Sheets[4];
        }
        private void closeExcel()
        {
            // 保存
            this.mWorkbook.Save();

            // ブックをクローズする
            this.mWorkbook.Close();

            if (null != this.mExcel)
            {
                // 警告無視設定
                this.mExcel.DisplayAlerts = false;

                // アプリケーションの終了
                this.mExcel.Quit();

                // アプリケーションのオブジェクトの解放
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.mWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.mExcel);
            }
        }
        private void writeDataToExcel(double[] checkData)
        {
            double headMove3D = checkData[0];
            double middleMove3D = checkData[1];
            double difTime = checkData[2];
            double speedHead = checkData[3];
            double speedMiddle = checkData[4];

            // 値を「A1」に代入
            int useRow = this.mSheet.UsedRange.Rows.Count;
            useRow += 1;
            this.mSheet.Cells[useRow, 1].Value = headMove3D;
            this.mSheet.Cells[useRow, 2].Value = middleMove3D;
            this.mSheet.Cells[useRow, 3].Value = speedHead;
            this.mSheet.Cells[useRow, 4].Value = speedMiddle;
            this.mSheet.Cells[useRow, 5].Value = difTime;

            if (this.fallDown)
            {
                this.mSheet.Cells[useRow, 5].Value = "〇";
            }
            else
            {
                this.mSheet.Cells[useRow, 5].Value = "×";
            }



        }

    }
    public class CheckData
    {

        public double[] headXInfo;
        public double[] headYInfo;
        public double[] headZInfo;
        public double[] middleXInfo;
        public double[] middleYInfo;
        public double[] middleZInfo;
        public long[] time;
        public bool fallDown = false;

        public CheckData()
        {
            headXInfo = new double[] { };
            headYInfo = new double[] { };
            headZInfo = new double[] { };
            middleXInfo = new double[] { };
            middleYInfo = new double[] { };
            middleXInfo = new double[] { };
            middleZInfo = new double[] { };
            headZInfo = new double[] { };
            time = new long[] { };
        }

    }
}

    
