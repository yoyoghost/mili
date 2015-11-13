/**
 * Excel导出测试类
 * 
 * @author yoyo
 *
 */
public class TestBean {

	@FieldMeta(aliasName = "姓名")
	private String name;

	@FieldMeta(aliasName = "年龄")
	private int age;

	@FieldMeta(aliasName = "性别")
	private String sex;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public String getSex() {
		return sex;
	}

	public void setSex(String sex) {
		this.sex = sex;
	}

}
