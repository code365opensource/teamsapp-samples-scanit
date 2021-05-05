import React, { useEffect, useState } from "react";
import "./App.css";
// 导入UI组件
import {
  Text,
  Button,
  List,
  AcceptIcon,
  ApprovalsAppbarIcon,
  Flex,
  PopupIcon,
  Alert,
} from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";

/* 顶部的问候组件 */
function Greeting(props: { UserName?: string }) {
  return (
    <Text content={props.UserName + ", 你好！"} important size="largest" />
  );
}

/* 扫码开箱按钮 */
function ActionBar(props: { click: any }) {
  return (
    <Button
      content="扫码开箱"
      icon={<PopupIcon />}
      iconPosition="before"
      primary
      onClick={() => props.click()}
    />
  );
}

/* 底部的固定文字 */
function Footer() {
  return <Text align="center" content="版权所有@code365.xyz" />;
}

/* 用户的借用记录 */
function History(props: { Items?: IHistoryItem[]; click: any }) {
  if (props.Items === undefined || props.Items.length === 0) {
    return <></>;
  }

  let items = props.Items?.map((x) => {
    return {
      content: x.EndTime ? (
        "归还时间：" + x.EndTime
      ) : (
        <Button content="归还" onClick={() => props.click(x.BoxNumber)} />
      ),
      header: "箱号:" + x.BoxNumber.toString() + (x.EndTime ? "" : ",正在使用"),
      headerMedia: x.StartTime,
      key: x.BoxNumber.toString(),
      media: x.EndTime ? <AcceptIcon /> : <ApprovalsAppbarIcon />,
    };
  });

  return <List items={items} />;
}

// 如果当前没有可用箱子时显示的提示
function NotAvailable() {
  return <Text content="当前没有箱子可用，请稍后再试" />;
}

// 当前有可用的箱子时显示的提示和操作按钮
function Available(props: { id?: number; confirm: any }) {
  return (
    <Flex column>
      <Text content={"恭喜，当前" + props.id + "号箱子可用，请立即使用吧"} />
      <Button
        content="马上打开"
        primary
        onClick={() => props.confirm(props.id)}
      />
    </Flex>
  );
}
// 历史记录类型定义
interface IHistoryItem {
  StartTime?: string;
  EndTime?: string;
  BoxNumber: number;
}

// 应用的主组件
function App() {
  //定义几个状态
  const [state, setState] = useState<number>(0);
  //这个用来保存用户的名称
  const [userName, setUserName] = useState<string>();
  //这个用来保存用户的历史记录
  const [historyItems, setHistoryItems] = useState<IHistoryItem[]>();
  //这个用来记录当前用户所选择的箱子编号
  const [selectedBox, setSelectedBox] = useState<number>();
  //这个用来显示一些消息提醒
  const [message, setMessage] = useState<string>();

  //这个方法用来获取某个用户的历史记录
  const getHistoryItems = async (userName: string): Promise<IHistoryItem[]> => {
    // 此处作为演示目的，仅读取本地存储，所以其实用户名已经不重要。真实的开发中，可以把这个username的信息，传递给后台服务，进行搜索查询。
    const found = localStorage.getItem("history");
    return found ? JSON.parse(found) : [];
  };

  // 这个方法会通过 “扫码开箱” 这个按钮来调用
  const scan = () => {
    //设置一些参数，这里是指30秒超时，如果没有操作
    const config: microsoftTeams.media.BarCodeConfig = {
      timeOutIntervalInSec: 30,
    };

    microsoftTeams.media.scanBarCode(
      (error: microsoftTeams.SdkError, decodedText: string) => {
        if (error) {
          let errorMessage;
          switch (error.errorCode) {
            case 100:
              errorMessage = "平台不支持";
              break;
            case 500:
              errorMessage = "内部错误";
              break;
            case 1000:
              errorMessage = "权限不足，用户没有接受";
              break;
            case 3000:
              errorMessage = "硬件不支持";
              break;
            case 4000:
              errorMessage = "错误的参数";
              break;
            case 8000:
              errorMessage = "用户取消操作";
              break;
            case 8001:
              errorMessage = "用户操作超时";
              break;
            case 9000:
              errorMessage = "平台太老";
              break;
            default:
              errorMessage = "未知错误";
              break;
          }
          setMessage("发生错误:" + errorMessage);
        } else if (decodedText) {
          // 这里是指扫描到了有关的二维码，本例不做具体的检测。通常会是一个网址，然后这个网址跟具体某个柜子是有关系的，然后通过这个网址发送请求，完成后续的开箱的操作。
          // 本例，只要扫描到了，就访问本地存储去尝试进行登记。
          setMessage("扫码成功:" + decodedText);
          // 随机生成一个箱子编号，也可能没有箱子的情况
          let id = Math.floor(Math.random() * 30);
          let available = id % 2 === 0;
          if (available) {
            let items = historyItems;
            items?.push({
              BoxNumber: id,
            });

            localStorage.setItem("history", JSON.stringify(items));
            setSelectedBox(id);
            setState(4);
          } else {
            setState(3);
            setTimeout(() => {
              setState(1);
            }, 5000);
          }
        }
        //设置两秒后自动消息
        setTimeout(() => {
          setMessage("");
        }, 2000);
      },
      config
    );
  };
  //归还某个箱子
  const returnBox = (id: number) => {
    const found = localStorage.getItem("history");
    const items: IHistoryItem[] = found ? JSON.parse(found) : [];
    const item = items.find((x) => x.BoxNumber === id);
    const options: Intl.DateTimeFormatOptions = {
      dateStyle: "short",
      timeStyle: "short",
    };

    const fmt = new Intl.DateTimeFormat("zh-CN", options);
    if (item) {
      item.EndTime = fmt.format(new Date());
    }
    localStorage.setItem("history", JSON.stringify(items));
    setHistoryItems(items);
    setState(1);
  };
  //确认借用某个箱子
  const confirm = (id: number) => {
    const found = localStorage.getItem("history");
    const items: IHistoryItem[] = found ? JSON.parse(found) : [];
    const item = items.find((x) => x.BoxNumber === id);
    const options: Intl.DateTimeFormatOptions = {
      dateStyle: "short",
      timeStyle: "short",
    };

    const fmt = new Intl.DateTimeFormat("zh-CN", options);
    if (item) {
      item.StartTime = fmt.format(new Date());
    }
    localStorage.setItem("history", JSON.stringify(items));
    setHistoryItems(items);
    setState(2);
  };

  // 这个只在初始化时调用，这里是读取用户信息和他历史记录信息
  useEffect(() => {
    localStorage.clear();
    // 初始化Teams，这个是必须调用的
    microsoftTeams.initialize();
    // 获取当前运行的环境上下文
    microsoftTeams.getContext(async (context) => {
      //告诉Teams当前应用已经初始化成功，通常会把“正在加载”的图标关闭掉
      microsoftTeams.appInitialization.notifySuccess();
      //获取用户的信息，为保护隐私，这里只能获取到userPrincipalName，通常就是用户的唯一的邮箱
      const user = context.userPrincipalName;
      // 设置当前的用户信息，以便通知界面更新
      setUserName(user);
      if (user) {
        // 获取当前用户的历史记录
        const items = await getHistoryItems(user);
        // 设置当前用户的历史记录，以便通知界面更新
        setHistoryItems(items);
        //检查是否有未归还的记录，如果有，则状态设置为2，否则为1
        let found = items.find((x) => x.EndTime === undefined);
        // 设置当前应用的状态
        setState(found ? 2 : 1);
      }
    });
  }, []);
  // 根据此前的设计，根据不同的状态，切换不同的组件，组成最合适的界面
  return (
    <Flex column>
      <Greeting UserName={userName} />
      {message && message?.length > 0 && <Alert content={message}></Alert>}
      {state === 0 && <Text content="这个网页只能在Teams中运行..." />}
      {state === 1 && <ActionBar click={scan} />}
      {state === 3 && <NotAvailable />}
      {state === 4 && <Available id={selectedBox} confirm={confirm} />}
      <History Items={historyItems} click={returnBox} />
      <Footer />
    </Flex>
  );
}
// 导出这个应用，传递给index.tsx
export default App;
