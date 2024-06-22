#include <napi.h>

class Utility {
 public:
  static Napi::Object Init(Napi::Env env, Napi::Object exports);
  static Napi::Array Cell(const Napi::CallbackInfo& info);
};
