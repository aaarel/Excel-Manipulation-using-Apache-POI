package Model;

import java.util.Objects;

/**
 * Created by ARIELPE on 4/2/2017.
 */
public class Customer {
	private String id;
	private String name;
	private String type;


	public Customer(String id) {
		this.id = id;
	}

	public Customer(String id, String name) {
		this.id = id;
		this.name = name;
	}

	public Customer(String id, String name, String type) {
		this.id = id;
		this.name = name;
		this.type = type;
	}


	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		Customer customer = (Customer) o;
		return Objects.equals(id, customer.id);
	}

	@Override
	public int hashCode() {
		return Objects.hash(id);
	}
}
