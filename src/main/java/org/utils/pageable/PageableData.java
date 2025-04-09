package org.utils.pageable;

import java.util.List;

@FunctionalInterface
public interface PageableData<T> {


    List<T> getData(Integer page, Integer size);

}
