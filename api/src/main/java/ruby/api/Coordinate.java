package ruby.api;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.Id;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

import static lombok.AccessLevel.PROTECTED;

@Entity
@Getter
@NoArgsConstructor(access = PROTECTED)
public class Coordinate {

    @Id
    @GeneratedValue
    private Long id;

    private String nodeId;
    private String arsId;
    private String stationName;
    private Double longitude;
    private Double latitude;

    @Builder
    public Coordinate(String nodeId, String arsId, String stationName, Double longitude, Double latitude) {
        this.nodeId = nodeId;
        this.arsId = arsId;
        this.stationName = stationName;
        this.longitude = longitude;
        this.latitude = latitude;
    }
}
