#include <napi.h>
#include <xlsxwriter.h>

class Format : public Napi::ObjectWrap<Format> {
 public:
  static Napi::Object Init(Napi::Env env, Napi::Object exports);
  static Napi::Value NewInstance(Napi::Env env, lxw_format* format);
  static lxw_format* Get(Napi::Value value);
  Format(const Napi::CallbackInfo& info);

 private:
  Napi::Value SetAlign(const Napi::CallbackInfo& info);
  Napi::Value SetBgColor(const Napi::CallbackInfo& info);
  Napi::Value SetFgColor(const Napi::CallbackInfo& info);
  Napi::Value SetBorderColor(const Napi::CallbackInfo& info);
  Napi::Value SetBottomColor(const Napi::CallbackInfo& info);
  Napi::Value SetTopColor(const Napi::CallbackInfo& info);
  Napi::Value SetLeftColor(const Napi::CallbackInfo& info);
  Napi::Value SetRightColor(const Napi::CallbackInfo& info);
  Napi::Value SetFontColor(const Napi::CallbackInfo& info);
  Napi::Value SetFontName(const Napi::CallbackInfo& info);
  Napi::Value SetFontScript(const Napi::CallbackInfo& info);
  Napi::Value SetFontSize(const Napi::CallbackInfo& info);
  Napi::Value SetFontStrikeout(const Napi::CallbackInfo& info);
  Napi::Value SetBold(const Napi::CallbackInfo& info);
  Napi::Value SetItalic(const Napi::CallbackInfo& info);
  Napi::Value SetUnderline(const Napi::CallbackInfo& info);
  Napi::Value SetBorder(const Napi::CallbackInfo& info);
  Napi::Value SetBottom(const Napi::CallbackInfo& info);
  Napi::Value SetTop(const Napi::CallbackInfo& info);
  Napi::Value SetLeft(const Napi::CallbackInfo& info);
  Napi::Value SetRight(const Napi::CallbackInfo& info);
  Napi::Value SetNumFormat(const Napi::CallbackInfo& info);
  lxw_format* format = nullptr;
};
