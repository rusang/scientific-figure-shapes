<!-- Slide number: 1 -->

MPHF-Net 工业表面缺陷检测网络结构设计图

输入图像
Backbone: CSP-MEEM
Neck: FPN-PAN （LOSC + CGAFusion）
图例说明
Head: RT-DETR 原始检测头
多尺度细粒度特征提取

P5 / C5 （20×C）

Top-Down 路径
Bottom-Up 路径

P4 / C4 （40×C）

P5
C5

IoU-Aware

![01_R1_pcb.png](SUMMER_R1_pcb.jpg)

+

20×C
Query Selection
P3 / C3 （80×C）

UP
↑
C5
····

Down
↓
20×C

UP
↑
上采样 （× 2）

+

C4
CGAFusion

Down
↓
下采样 （× 2）

LOSC

C4
Transformer Decoder × 6 层

P4

+
40×C

+
拼接 （Concat）

UP
↑

40×C

方向性下采样建模

LOSC

••••

Down
↓

C3

+
高低频自适应融合

CGAFusion

CGAFusion
C3
80×C
640 × 640 × 3

P3

输出：
类别概率 + 归一化边界框坐标

+

⋮
80×C
数据流向

跨层连接

输出多尺度特征

CGAFusion（高低频自适应融合）
LOSC（方向性下采样建模）

模块一：CSP-MEEM（多尺度细粒度特征提取）
模块二：LOSC（方向性下采样建模）
模块三：CGAFusion（高低频自适应融合）
输入图像
输入特征
输入特征 Fh（高频）
•  多尺度并行感知，捕获不同
大小的缺陷信息

CSP
分割
•  分离高频细节与低频结构信息
•  引入方向性卷积核
高频分支
（细节信息）
•  沿关键方向提取特征
•  增强细节特征表达，保留
边缘和纹理
•  空间联合注意力自适应融合

···
分支1
分支2
分支n
输入特征 Fl（低频）

Down
↓

•  抑制无关干扰，保留方向信息

+
低频分支
（结构信息）

融合
•  突出关键细节，抑制噪声干扰
•  EMA 注意力增强关键特征，
抑制冗余信息
•  增强对工业表面纹理与缺陷
形态的建模能力

EMA 注意力

•  提升特征表达与泛化能力
输出特征
输出特征
输出特征

三模块协同优势
输入到 RT-DETR Head
输出多尺度特征金字塔
检测结果

![02_adv_precision.png](SUMMER_E_b4_icon0.jpg)
Transformer Decoder × 6 层
IoU-Aware
Query Selection
更精细：
保留小目标细粒度特征

![03_adv_robust.png](SUMMER_E_b4_icon1.jpg)

····
更鲁棒：
增强方向性缺陷感知

····

P3
（80×C）
P4
（40×C）
P5
（20×C）

![04_adv_efficient.png](SUMMER_E_b4_icon2.jpg)
类别概率
归一化边界框坐标
更高效：
自适应融合多尺度特征

<!-- Slide number: 2 -->

![mphf-net-architecture.png](SUMMER_R_reference_full.jpg)
