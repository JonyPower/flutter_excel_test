import 'dart:io';

import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:path/path.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Excel Test',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: const MyHomePage(title: 'Excel Package Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});
  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  final path = 'c:\\excel_test\\test.xlsx';

  void test1() {
    try {
      print('start teste1 - red color');
      var bytes = File(path).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      excel.updateCell(
          'Plan1', CellIndex.indexByString('A1'), 'Red Color Expected',
          cellStyle: CellStyle(
            backgroundColorHex: '#FF0000',
          ));

      List<int>? fileBytes = excel.save();
      File(join(path))
        ..createSync(recursive: false)
        ..writeAsBytesSync(fileBytes!);

      print('save teste1 ok');
    } catch (e) {
      print('Error on test1');
      print(e);
    }
  }

  void test2() {
    try {
      print('start teste2 - yellow color');
      var bytes = File(path).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      excel.updateCell(
          'Plan1', CellIndex.indexByString('A2'), 'Yellow Color Expected',
          cellStyle: CellStyle(
            backgroundColorHex: '#FFFF00',
          ));

      List<int>? fileBytes = excel.save();
      File(join(path))
        ..createSync(recursive: false)
        ..writeAsBytesSync(fileBytes!);

      print('save teste2 ok');
    } catch (e) {
      print('Error on test2');
      print(e);
    }
  }

  void test3() {
    try {
      print('start teste3 - green color');
      var bytes = File(path).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      excel.updateCell(
          'Plan1', CellIndex.indexByString('A3'), 'Green Color Expected',
          cellStyle: CellStyle(
            backgroundColorHex: '#09FF32',
          ));

      List<int>? fileBytes = excel.save();
      File(join(path))
        ..createSync(recursive: false)
        ..writeAsBytesSync(fileBytes!);

      print('save teste3 ok');
    } catch (e) {
      print('Error on test3');
      print(e);
    }
  }


  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text(widget.title),
      ),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: <Widget>[
            TextButton(
                onPressed: () {
                  test1();
                },
                child: const Text('Test1')),
            TextButton(
                onPressed: () {
                  test2();
                },
                child: const Text('Test2')),
            TextButton(
                onPressed: () {
                  test3();
                },
                child: const Text('Test3'))
          ],
        ),
      ),
    );
  }
}
