//// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
///		
/// 				2019秋チラシ　オモテ面用スクリプト
///							100点満点 x 新規開校
/// 	
/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ


(function() {
///  ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
///   CSVConverterのコンストラクタ
///　
	var CSVConverter = function() {
		this.textWords = null;	// csvファイル1行目の項目名
		this.lines = [];
		this.textGroups = []; // 変数内のテキストを格納
		this.pathGroups = []; // 変数内のパスアイテムを格納
		this.doc = app.activeDocument;
		this.variables = app.activeDocument.variables;
		this.path = String(app.documents[0].fullName).replace(app.documents[0].name, "");
		
		//CSVの何列目に画像ファイルの項目があるか
		this.imageIndex = 13; // 13列目
		
		for (var i = 0, n = this.variables.length; i < n; i++) {
			var group = this.variables[i].pageItems[0]; // イラレで設定した変数の取得
			
			// 変数内のテキストボックスの取得
			if (group.textFrames.length != 0) {
				this.textGroups.push(group);
			}
			
		}

		// 変数内のパスアイテムの取得
		for (i = 0, n = this.variables.length; i < n; i++) {
			group = this.variables[i].pageItems[0];
			if (group.pathItems.length != 0) {
				this.pathGroups.push(group);
			}
		}
				
	};
	///  ｰｰｰｰｰｰｰCSVConverterのコンストラクタｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ　
	
	///  ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
	///   CSVConverterの各関数定義
	///  ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
	///  read 関数：CSV読み込み
		CSVConverter.prototype = {
		read: function() {
			var path = File.openDialog("CSVファイルを選択してください。");

			if (! path) {
				return;
			}

			var csv = new File(path);

			if (! csv.open("r", "", "")) {
				return;
			}

			this.lines = [];
			var text = csv.read(); // 選択したCSVファイルのすべてのテキストを読み込み
			var lines = text.split(String.fromCharCode(10)); // 改行コードでテキストを分割
				
			//$.writeln(lines); 
			
			// 1行ずつ、1セルずつの情報に分割
			for (var i = 0, n = lines.length; i < n; i++) {
				var line = lines[i];

				if (! line) {
					continue;
				}

				if (i == 0) {
					this.textWords = line.split(","); // 1行目は、項目名
				} else {
					this.lines.push(line.split(",")); // 2行目以降は、内容を取得
				}
			}
			//$.writeln(this.lines);

			csv.close();
		}, 		
		/// read関数
		///ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
		
		/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
		/// writeText関数
		writeText: function(i) {
			
				var textRef,textRef1,textRef2;
				
				try {
				
				// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
				// 教室名
				if(this.lines[i][1] == "") // 教室名２が空欄かどうか
				{
					// 教室名１のみの場合
					// textFramesを追加しちゃう
					textRef = doc.textFrames.add();
					textRef.contents =  this.lines[i][0]; 
					textRef.top = 260;
					textRef.left = 255;
					textRef.textRange.characterAttributes.size = 65;
					
					redraw();
										
					$.writeln("幅-前")
					$.writeln(String(textRef.width))
					
					// テキスト幅調整
					while(textRef.width > 80)
					{
						// フォントサイズを2ptずつ小さくする
						textRef.textRange.characterAttributes.size = textRef.textRange.characterAttributes.size - 2;
						redraw();
					}
					$.writeln("幅-あと")
					$.writeln(String(textRef.width))				
				}
				else
				{
					// 教室名2もある場合
					// font size 32pt
					// textFramesを追加しちゃう
					textRef1 = doc.textFrames.add();
					textRef1.contents =  this.lines[i][0]; 
					textRef1.top = 260;
					textRef1.left = 255;
					textRef1.textRange.characterAttributes.size = 32;
					
					textRef2 = doc.textFrames.add();
					textRef2.contents =  this.lines[i][1]; 
					textRef2.top = 268;
					textRef2.left = 255;
					textRef2.textRange.characterAttributes.size = 32;
					
					redraw();
										
					// テキスト幅調整
					while(textRef2.width > 80)
					{
						// フォントサイズを2ptずつ小さくする
						textRef2.textRange.characterAttributes.size = textRef2.textRange.characterAttributes.size - 2;
						redraw();
					}
					textRef1.textRange.characterAttributes.size = textRef2.textRange.characterAttributes.size;
					redraw();
									
				}
				// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
				} catch (e) {}		
				
				// 保存
				this.exportJpeg();
				
				// 削除
				if(typeof textRef !== "undefined"){
					textRef.remove();
				}
				if(typeof textRef1 !== "undefined"){
					textRef1.remove();
				}
				if(typeof textRef2 !== "undefined"){
					textRef2.remove();
				}
				
					

		},
		/// writeText関数
		/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
		
		/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
		///　writeImage
		writeImage: function(i) {
				
				
				var groups = this.pathGroups[0];
				var paths = [];

				if (groups.pathItems.length != 0) {
					paths = groups.pathItems;
					$.writeln("画像はあるよ" )
				}
				
				var pItem1,mask1,holder1;
				var pItem2,mask2,holder2;

					/*	
					// qrコード
					try {
						var rect = paths[0];
						var file = new File(this.path +"img/qr/" +this.lines[i][this.imageIndex - 1]);
						pItem1 = activeDocument.placedItems.add();
						pItem1.file = file;
						this._createPosition(pItem1, rect);
						mask1 = activeDocument.pathItems.rectangle(rect.top, rect.left, rect.width, rect.height);
						mask1.stroke = true;
						mask1.filled = true;
						holder1 = app.activeDocument.groupItems.add();
						pItem1.move(holder1, ElementPlacement.PLACEATEND);
						mask1.move(holder1, ElementPlacement.PLACEATBEGINNING);
						holder1.clipped = true;
					} catch (e) {}
					*/
					
					
					/*
					// score
					try {
						var rect = paths[1];
						var fname = "";
						if(this.lines[i][7] == "1"){
							fname = "score100.png";
						}							
						else{
							fname = "score50.png";
						}
						var file = new File(this.path +"/score/" +fname);
						var pItem = activeDocument.placedItems.add();
						pItem.file = file;
						this._createPosition(pItem, rect);
						var mask = activeDocument.pathItems.rectangle(rect.top, rect.left, rect.width, rect.height);
						mask.stroke = true;
						mask.filled = true;
						var holder = app.activeDocument.groupItems.add();
						pItem.move(holder, ElementPlacement.PLACEATEND);
						mask.move(holder, ElementPlacement.PLACEATBEGINNING);
						holder.clipped = true;
					} catch (e) {}
					*/
					
					
   					//	 追加画像の削除
   					//pItem1.remove(); mask1.remove(); holder1.remove(); 
					

				//}
			//}
		},
				
		/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
		/// _createPosition
		_createPosition: function(targetA, targetB) {
			var widthA = targetA.width;
			var widthB = targetB.width;
			var heightA = targetA.height;
			var heightB = targetB.height;
			targetA.left = targetB.left;
			targetA.top = targetB.top;

			if (widthA > widthB) {
				targetA.left = targetA.left - ((widthA - widthB) / 2);
			}

			if (heightA > heightB) {
				targetA.top = targetA.top + ((heightA - heightB) / 2);
			}
		},
		/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
		
		/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
		///　exportJpeg関数
		exportJpeg: function() {
		// 保存
		// 保存ファイルの詳細
		var exportOpt = new ExportOptionsJPEG();
		var type = ExportType.JPEG;
   			 		
   		converter.path = String(app.documents[0].fullName).replace(app.documents[0].name, "");
   		var fname =   converter.path  + "img/rslt/" +this.lines[i][0] + ".jpg";
   		var fileSpec = new File(fname);
   		exportOpt.antiAliasing = false;
   		exportOpt.qualitySetting = 70;
   		// 保存
   		app.activeDocument.exportFile(fileSpec, type, exportOpt);   					
	}
	/// exportJpeg関数
	/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ

		
	};
	/// writeImage関数
	/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
	
	
	
	/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ
	/// 実際の処理
	/// ｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰｰ

	$.writeln("Start");
	var converter = new CSVConverter();
	converter.read();
	
	// ★宿題：チラシタイプと点数によるエラー処理作成
	
	for (var i = 0, n = converter.lines.length; i < n; i++) {
		converter.writeText(i);
		//converter.writeImage(i);   				
   		}
   		// 元に戻す
   		/*
   		var textFrames = converter.textGroups[0].textFrames;				
		for (var j = 0, o = textFrames.length; j < o; j++) {
			try {
				var textFrame = textFrames[j];
				textFrame.contents = converter.textWords[j];
				}
					catch (e) {}
				}
				*/
})();
