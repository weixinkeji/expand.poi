package weixinkeji.vip.expand.poi;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

@Retention(RUNTIME)
@Target(FIELD)
public @interface JWEOffice {
	
	public String value() default "";
	public String title() default "";
	public int sort() default 0;
	public String dateformat() default "yyyy-MM-dd";
	
}
