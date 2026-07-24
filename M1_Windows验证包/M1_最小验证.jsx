var m1Args = (typeof arguments !== "undefined") ? arguments : [];
if (m1Args.length < 1 || !m1Args[0]) {
    throw new Error("M1 JSX 缺少 marker 输出路径参数 arguments[0]。");
}

var m1MarkerPath = String(m1Args[0]);
(function (markerPath) {
    var file = new File(markerPath);
    file.encoding = "UTF-8";
    if (!file.open("w")) {
        throw new Error("M1 JSX 无法写入 marker 文件：" + markerPath);
    }

    file.writeln("M1_RESULT|成功|Photoshop JSX 已执行|" + (new Date()).getTime());
    file.close();
    return "M1_RETURN|成功|" + markerPath;
})(m1MarkerPath);
