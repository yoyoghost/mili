import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
/**
 * Excel 导出注解类
 * @author yoyo
 *
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface FieldMeta {

	/**
	 * 字段别名
	 * 
	 * @return
	 */
	String aliasName() default "";


	/**
	 * 字段描述
	 * 
	 * @return
	 */
	String description() default "";
	
	/**
	 * 字段备注
	 * 
	 * @return
	 */
	String remark() default "";
}
