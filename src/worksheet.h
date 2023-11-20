#include <napi.h>
#include <xlsxwriter.h>

class Worksheet : public Napi::ObjectWrap<Worksheet> {
 public:
  static Napi::Object Init(Napi::Env env, Napi::Object exports);
  static Napi::Value NewInstance(Napi::Env env, lxw_worksheet* worksheet);
  Worksheet(const Napi::CallbackInfo& info);

 private:
  Napi::Value InsertChart(const Napi::CallbackInfo& info);
  Napi::Value InsertImage(const Napi::CallbackInfo& info);
  Napi::Value MergeRange(const Napi::CallbackInfo& info);
  Napi::Value SetColumn(const Napi::CallbackInfo& info);
  Napi::Value SetRow(const Napi::CallbackInfo& info);
  Napi::Value SetFooter(const Napi::CallbackInfo& info);
  Napi::Value SetHeader(const Napi::CallbackInfo& info);
  Napi::Value WriteBoolean(const Napi::CallbackInfo& info);
  Napi::Value WriteDatetime(const Napi::CallbackInfo& info);
  Napi::Value WriteFormula(const Napi::CallbackInfo& info);
  Napi::Value WriteNumber(const Napi::CallbackInfo& info);
  Napi::Value WriteString(const Napi::CallbackInfo& info);
  lxw_worksheet* worksheet = nullptr;
};
