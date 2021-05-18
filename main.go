package main

import (
	"log"
	"net/http"
	"office-utils/controller"
)



func main() {
	http.HandleFunc("/build_excel", Cors(controller.BuildExcel))
	err := http.ListenAndServe(":8086", nil)
	if err != nil {
		log.Fatal("ListenAndServe: ", err)
	}
}



func Cors(f http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")  // 允许访问所有域，可以换成具体url，注意仅具体url才能带cookie信息
		w.Header().Add("Access-Control-Allow-Headers", "Content-Type,Authorization") //header的类型
		w.Header().Add("Access-Control-Allow-Credentials", "true") //设置为true，允许ajax异步请求带cookie信息
		w.Header().Add("Access-Control-Allow-Methods", "*") //允许请求方法

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusNoContent)
			return
		}
		f(w, r)
	}
}