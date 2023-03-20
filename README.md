HwpApi
================

<!-- WARNING: THIS FILE WAS AUTOGENERATED! DO NOT EDIT! -->

This file will become your README and also the index of your
documentation.

## Install

``` sh
pip install HwpApi
```

## How to use

Fill me in please! Don’t forget code examples:

## 왜 HwpApi를 만들었나요?

가장 큰 이유는 스스로 사용하기 위해서 입니다. 직장인으로 많은 한글
문서를 편집하고 작성하곤 하는데 단순 반복업무가 너무 많다는 것이
불만이었습니다. 이런 문제를 해결하는 방법으로 한글 자동화에 대한
이야기를 파이콘에서 보게 되었습니다. 특히 ‘회사원 코딩’ 님의 블로그와
영상이 많은 참조가 되었습니다.

다만 그 과정에서 설명자료가 부족하기도 하고 예전에 작성했던 코드들을
자꾸 찾아보게 되면서 아래아 한글 용 파이썬 패키지가 있으면 좋겠다는
생각을 했습니다. 특히 업무를 하면서 엑셀 자동화를 위해 xlwings를 사용해
보면서 파이썬으로 사용하기 쉽게 만든 라이브러리가 코딩 작업 효율을 엄청
올린다는 것을 깨닫게 되었습니다.

제출 마감까지 해야 할 일들을 빠르게 하기 위해서 빠르게 한글 자동화가
된다면 좋겠다는 생각으로 만들게 되었습니다.

기본적인 철학은 xlwings을 따라하고 있습니다. 기본적으로는 자주 쓰이는
항목들을 사용하기 쉽게 정리한 메소드 등으로 구현하고, 부족한 부분은
`App.api`형태로 `win32com`으로 하는 것과 동일한 작업이 가능하게 하여
한글 api의 모든 기능을 사용할 수 있도록 구현하였습니다.

메소드로 만드는 것에는 아직 고민이 있습니다. chain과 같은 형태로
여러가지 콤비네이션을 사전에 세팅을 해야 하나 싶은 부분도 있고 실제로
유용하게 사용할 수 있는 여러가지 아이템 등도 있어서 어떤 부분까지 이
패키지에 구현할지는 고민하고 있습니다.

다만 이런 형태의 작업을 통해서 어쩌면 hwp api wrapper가 활성화 되어서
단순 작업을 자동화 할 수 있기를 바라고 있습니다.

## 기존 코드와 연동성 비교하기

[회사원 코딩](https://employeecoding.tistory.com/72)에 가보시면 아래와
같이 자동화 코드가 있습니다.

``` python
import win32com.client as win32
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True

act = hwp.CreateAction("InsertText")
pset = act.CreateSet()
pset.SetItem("Text", "Hello\r\nWorld!")
act.Execute(pset)
```

이 코드는 기본적으로 verbose라고 볼 만한 상황입니다. 이 코드를
`HwpApi`를 사용하면 아래와 같이 간결하게 정리가 됨을 볼 수 있습니다.

``` python
from HwpApi.core import App

app = App()
action = app.actions.InsertText()
p = action.pset
p.Text = "Hello\r\nWorld!"
action.run()
```

    True
