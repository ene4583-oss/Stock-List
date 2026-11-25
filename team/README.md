# Smart Stock 多人版

员工和主管用管理员分配的用户名及密码登录。只有管理员需要 Google 账号。使用私人测试表格和单独的 Apps Script 项目验证后再上线。

1. 在测试表格打开「扩展程序 → Apps Script」。粘贴 [Code.gs](Code.gs)，新建名为 Index 的 HTML 文件后粘贴 [Index.html](Index.html)。
2. 左侧点「服务 → ＋ → Google Sheets API → 添加」。标准 Cloud 项目还需要在对应项目启用 Sheets API。
3. 保存并以表格拥有者身份运行 `setupDatabase`。如果表头不兼容，停止并备份，先处理迁移。初始化保留原有资料。
4. 以同一个账号部署 Web App，执行身份选「我」。访问权限选「任何人」（允许未登录 Google 的访客）。若组织禁止匿名 Web App，员工登录方案需要更换部署环境。表格本身继续保持私人共享。
5. 用表格拥有者的 Google 账号打开网页，进入 Manage，创建工厂、物品和分配关系。为员工选择身份、工厂和用户名。系统生成密码，只在开户或勾选「生成新密码」并保存后显示一次。请私下发给本人。
6. 用无痕窗口（未登录 Google）测试员工账号。确认只能看到分配的工厂、物品和本人领用历史。用主管账号再验证收货及所属工厂的库存权限。用管理员录入 Opening Stock 后才能测试扣库。

登录采用随机生成的密码，员工无需自行设定密码。密码以服务端密钥保护的验证值存储。网页不保存登录令牌到浏览器存储，刷新需重新登录。会话最长八小时，重新登录、重设密码、停用或更新用户会使旧会话失效。匿名访客可以打开登录页，业务方法要求通过后端身份及权限校验。仅将表格和脚本的编辑权限交给可信维护者。

管理员身份要求 Google 向部署返回拥有者的真实邮箱。请先实测拥有者登录及无 Google 账号的访客流程。Apps Script 的执行和 PropertiesService 配额适用于本版本，大规模同时使用前需做负载测试。

## Items and Suggested Buy

In **Manage → Items**, select the factories that use the item, enter an optional estimated quantity per day for each factory, and upload a photo if needed. Saving factory changes preserves existing stock balances and history. **User accounts** is at the bottom of Manage.

After updating an existing deployment, run `authorizeItemPhotos` once in the editor as the owner and approve Drive access before deploying the new version. Photos are resized in the browser and kept in a private Drive folder. The web app serves them only after checking account, item, and factory access.

**Home → Suggested Buy** uses recent observed consumption when at least seven days are available, or the per-factory estimate when there is less history. A day without a record is excluded unless coverage has been confirmed. Initial lead time, review period and safety defaults are 3, 7 and 3 days. Recommendations subtract stock and eligible outstanding orders due within the planning horizon, then round up to the item's pack size. The list does not place orders. The current screen reads eligible orders already in PURCHASE_ORDERS.
