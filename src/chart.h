#include <napi.h>
#include <xlsxwriter.h>

class Chart : public Napi::ObjectWrap<Chart> {
 public:
  static Napi::Object Init(Napi::Env env, Napi::Object exports);
  static Napi::Value NewInstance(Napi::Env env, lxw_chart* chart);
  static lxw_chart* Get(Napi::Value value);
  Chart(const Napi::CallbackInfo& info);

 private:
  Napi::Value AddSeries(const Napi::CallbackInfo& info);
  Napi::Value SetTitleName(const Napi::CallbackInfo& info);
  Napi::Value SetTitleNameFont(const Napi::CallbackInfo& info);
  lxw_chart* chart = nullptr;
};
