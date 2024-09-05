package ve.powercalc.entity;

import jakarta.persistence.*;
import lombok.Getter;

@Entity
@Table (name="power_count_group")
@Getter
public class Group {

    @Id
    @Column (name = "id")
    @GeneratedValue (strategy = GenerationType.IDENTITY)
    public Integer id;

    @Column (name = "idname")
    public String name;
}
