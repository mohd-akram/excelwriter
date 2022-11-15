#include <napi.h>
#include <xlsxwriter.h>

class Format : public Napi::ObjectWrap<Format> {
 public:
  static Napi::Object Init(Napi::Env env, Napi::Object exports);
  static Napi::Value NewInstance(Napi::Env env, lxw_format* format);
  static lxw_format* Get(Napi::Value value);
  Format(const Napi::CallbackInfo& info);

 private:
  Napi::Value SetBold(const Napi::CallbackInfo& info);
  Napi::Value SetNumFormat(const Napi::CallbackInfo& info);
  lxw_format* format = nullptr;
};
