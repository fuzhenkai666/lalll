/*
 * @Author: 杨仕明 shiming.y@qq.com
 * @Date: 2024-11-06 15:35:55
 * @LastEditors: 杨仕明 shiming.y@qq.com
 * @LastEditTime: 2024-11-06 18:30:46
 * @FilePath: /next_word_auto/src/app/page.tsx
 * @Description: 
 * 
 * Copyright (c) 2024 by ${git_name_email}, All Rights Reserved. 
 */


'use client';


import { useEffect } from "react";
import { useRouter } from "next/navigation";
import HomeContent from "../components/homeContent";

export default function Home() {
  const router = useRouter();

  useEffect(() => {
    const isAuthenticated = localStorage.getItem("isAuthenticated");
    if (isAuthenticated !== "true") {
      router.push("/password"); // 如果没有验证通过，跳转到密码页面
    }
  }, [router]);

  return (
    <HomeContent />  // 将主页面内容分离到另一个组件中
  );
}