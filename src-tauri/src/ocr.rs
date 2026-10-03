//! 调用 Windows 系统自带的 OCR（Windows.Media.Ocr）：离线运行、不增加安装包体积。
//! 输入一张 PNG，输出按行分组的单词及其像素坐标，由前端换算成和 PDF 文字层相同的格式。

use serde::Serialize;

#[derive(Serialize)]
pub struct OcrWord {
    pub text: String,
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

#[derive(Serialize)]
pub struct OcrOutput {
    /// 实际使用的识别语言，例如 "zh-Hans-CN"
    pub language: String,
    pub lines: Vec<Vec<OcrWord>>,
}

#[cfg(windows)]
pub fn recognize(png: &[u8]) -> Result<OcrOutput, String> {
    use windows::core::HSTRING;
    use windows::Globalization::Language;
    use windows::Graphics::Imaging::BitmapDecoder;
    use windows::Media::Ocr::OcrEngine;
    use windows::Storage::Streams::{DataWriter, InMemoryRandomAccessStream};
    use windows::Win32::System::WinRT::{RoInitialize, RO_INIT_MULTITHREADED};

    let err = |e: windows::core::Error| format!("OCR 失败：{}", e.message());

    // 后台线程需要先初始化 WinRT；已初始化时返回的错误可以忽略
    unsafe {
        let _ = RoInitialize(RO_INIT_MULTITHREADED);
    }

    // 优先用简体中文识别器，没有安装时退回系统界面语言
    let engine = ["zh-Hans-CN", "zh-Hans", "zh-CN"]
        .iter()
        .filter_map(|tag| Language::CreateLanguage(&HSTRING::from(*tag)).ok())
        .find(|lang| OcrEngine::IsLanguageSupported(lang).unwrap_or(false))
        .and_then(|lang| OcrEngine::TryCreateFromLanguage(&lang).ok())
        .or_else(|| OcrEngine::TryCreateFromUserProfileLanguages().ok())
        .ok_or("这台电脑没有可用的 Windows OCR 识别语言")?;

    let stream = InMemoryRandomAccessStream::new().map_err(err)?;
    let writer = DataWriter::CreateDataWriter(&stream).map_err(err)?;
    writer.WriteBytes(png).map_err(err)?;
    writer.StoreAsync().map_err(err)?.join().map_err(err)?;
    writer.FlushAsync().map_err(err)?.join().map_err(err)?;
    writer.DetachStream().map_err(err)?;
    stream.Seek(0).map_err(err)?;

    let decoder = BitmapDecoder::CreateAsync(&stream).map_err(err)?.join().map_err(err)?;
    let bitmap = decoder.GetSoftwareBitmapAsync().map_err(err)?.join().map_err(err)?;
    let result = engine.RecognizeAsync(&bitmap).map_err(err)?.join().map_err(err)?;

    let mut lines = Vec::new();
    for line in result.Lines().map_err(err)? {
        let mut words = Vec::new();
        for word in line.Words().map_err(err)? {
            let r = word.BoundingRect().map_err(err)?;
            words.push(OcrWord {
                text: word.Text().map_err(err)?.to_string(),
                x: r.X as f64,
                y: r.Y as f64,
                w: r.Width as f64,
                h: r.Height as f64,
            });
        }
        if !words.is_empty() {
            lines.push(words);
        }
    }
    let language = engine
        .RecognizerLanguage()
        .and_then(|l| l.LanguageTag())
        .map(|t| t.to_string())
        .unwrap_or_default();
    Ok(OcrOutput { language, lines })
}

#[cfg(not(windows))]
pub fn recognize(_png: &[u8]) -> Result<OcrOutput, String> {
    Err("OCR 目前只支持 Windows".into())
}
