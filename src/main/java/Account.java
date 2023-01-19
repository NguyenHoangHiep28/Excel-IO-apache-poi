import java.util.Date;

public class Account {
    private Integer id;
    private String name;
    private Double balance;
    private Date effectiveStartDate;
    public Account(){

    }
    public Account(Integer id, String name, Double balance, Date effectiveStartDate) {
        this.id = id;
        this.name = name;
        this.balance = balance;
        this.effectiveStartDate = effectiveStartDate;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Double getBalance() {
        return balance;
    }

    public void setBalance(Double balance) {
        this.balance = balance;
    }

    public Date getEffectiveStartDate() {
        return effectiveStartDate;
    }

    public void setEffectiveStartDate(Date effectiveStartDate) {
        this.effectiveStartDate = effectiveStartDate;
    }

    @Override
    public String toString() {
        return "Account{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", balance=" + balance +
                '}';
    }
}
