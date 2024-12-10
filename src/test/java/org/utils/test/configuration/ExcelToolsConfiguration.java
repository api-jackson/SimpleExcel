package org.utils.test.configuration;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.utils.extra.BooleanToYNString;

/**
 * @author Jackson
 * @date 2024/11/30
 * @description
 */
@Configuration
public class ExcelToolsConfiguration {


    @Bean(name = "booleanToYNString")
    public BooleanToYNString booleanToYNString() {
        return new BooleanToYNString();
    }

}
