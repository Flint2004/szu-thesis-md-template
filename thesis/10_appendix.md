# 附录 A 证明

本附录收录正文中两个主要定理的完整证明。

## A.1 定理 4.1（成功率下界）的完整证明

重述定理 4.1：在假设 4.1—4.3 下，冻结成功率 $p_{\mathrm{s}}$ 满足

$$
p_{\mathrm{s}} \ge 1 - \exp\left[-\frac{\theta\tau\cdot A_{\mathrm{eff}}(\alpha)\cdot r^{-\beta}}{Q^*}\right].
$$

**证明** 记青蛙在一次施术中所吸收的总负焓流为随机变量 $Q_{\mathrm{abs}}$。由假设 4.2 与 3.2 节推导，其期望为

$$
\mu := \mathbb{E}[Q_{\mathrm{abs}}] = \theta\tau\cdot A_{\mathrm{eff}}(\alpha)\cdot r^{-\beta}.
$$

冻结成功事件等价于 $Q_{\mathrm{abs}}\ge Q^*$。由于 $Q_{\mathrm{abs}}$ 是若干相互独立的表皮子区域冷气吸收量之和，根据 Poisson 型计数过程的经典结果，其分布可由参数为 $\mu$ 的 Poisson 分布近似。于是

$$
p_{\mathrm{s}} = \mathrm{Pr}[Q_{\mathrm{abs}}\ge Q^*] = 1 - \mathrm{Pr}[Q_{\mathrm{abs}} < Q^*].
$$

利用 Chernoff 型不等式的一侧结果，对 Poisson($\mu$) 而言当 $Q^*\le \mu$ 时

$$
\mathrm{Pr}[Q_{\mathrm{abs}}<Q^*] \le \exp\left[-\frac{(\mu - Q^*)^2}{2\mu}\right].
$$

在 $Q^*\ll \mu$ 的参数区间内，上式右端可进一步被 $\exp(-\mu/Q^*)$ 所主导（详见文献[7]附录推导）。由此

$$
p_{\mathrm{s}} \ge 1 - \exp\left[-\frac{\mu}{Q^*}\right] = 1 - \exp\left[-\frac{\theta\tau\cdot A_{\mathrm{eff}}(\alpha)\cdot r^{-\beta}}{Q^*}\right].
$$

证毕。

## A.2 定理 4.5（等边际妖力回报）的完整证明

重述定理 4.5：若 $p_{\mathrm{s}}$ 关于乘积 $\theta_k\tau_k$ 连续可微、单调递增且凹，则多目标并行冻结问题的最优解 $\{(\theta_k^*,\tau_k^*)\}$ 满足

$$
\frac{\partial p_{\mathrm{s}}}{\partial (\theta_k\tau_k)}\bigg|_{\theta_k^*,\tau_k^*} = \lambda,\quad \forall k,
$$

且约束等号成立。

**证明** 记 $u_k := \theta_k\tau_k$，则问题改写为

$$
\max_{\{u_k\ge 0\}}\sum_{k=1}^{K} g_k(u_k)\quad\mathrm{s.t.}\quad \sum_{k=1}^{K}\eta u_k \le E_{\max},
$$

其中 $g_k(u_k) := p_{\mathrm{s}}(u_k;\alpha_k,r_k)$。由假设 $g_k$ 连续可微、单调递增且凹。

拉格朗日函数为

$$
\mathcal{L}(\{u_k\},\lambda) = \sum_{k=1}^{K}g_k(u_k) - \lambda\left[\sum_{k=1}^{K}\eta u_k - E_{\max}\right].
$$

其 KKT 条件为：对每个 $k$，

$$
g_k'(u_k^*) - \lambda\eta \le 0,\quad u_k^*\ge 0,\quad u_k^*\cdot[g_k'(u_k^*)-\lambda\eta]=0,
$$

以及原问题约束 $\sum_k\eta u_k^*\le E_{\max}$ 与对偶互补 $\lambda\ge 0$。

由于 $g_k$ 单调递增、$g_k'>0$，且 $E_{\max}$ 有限，故在最优解处不可能有 $u_k^*=0$ 对所有 $k$ 同时成立。假设某 $u_k^*>0$，则互补松弛给出 $g_k'(u_k^*) = \lambda\eta$。又由 $g_k$ 的凹性，$g_k'(u_k)$ 关于 $u_k$ 单调递减，这意味着 $\lambda\eta$ 一旦确定，所有 $u_k^*$ 均由方程 $g_k'(u_k^*)=\lambda\eta$ 唯一决定。

最后论证约束取等。假设 $\sum_k\eta u_k^* < E_{\max}$，则可以把任意一个 $u_j^*$ 增加一点 $\delta>0$，由于 $g_j'(u_j^*)>0$，$\sum_k g_k$ 会严格增大，与 $u_k^*$ 是最优解矛盾。故 $\sum_k \eta u_k^* = E_{\max}$。

综上，对任意最优 $u_k^*>0$ 有 $g_k'(u_k^*) = \lambda\eta$，即等边际妖力回报。令 $\lambda' := \lambda\eta$ 可把 $\eta$ 吸收进拉格朗日乘子中，恢复正文中的记号。证毕。

## A.3 补充说明

本附录所用两不等式（Chernoff 型、Markov 型）的展开细节可参阅[7]第 3 章。对于更精细的成功率刻画（例如包含 $\alpha$ 漂移的情形），需要引入二阶矩估计，这已超出本文范围。
