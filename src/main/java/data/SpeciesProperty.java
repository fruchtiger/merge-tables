package data;

public class SpeciesProperty {

    public SpeciesProperty(String species, String property, Double doubleValue, String stringValue, Integer reference) {
	super();
	this.species = species;
	this.property = property;
	this.doubleValue = doubleValue;
	this.stringValue = stringValue;
	this.reference = reference;
    }

    public final String species;
    public final String property;
    public final String stringValue;
    public final Double doubleValue;
    public final Integer reference;

    @Override
    public String toString() {
	return "SpeciesProperty [species=" + species + ", property=" + property + ", stringValue=" + stringValue
		+ ", doubleValue=" + doubleValue + ", reference=" + reference + "]";
    }

    // @Override
    // public int hashCode() {
    // final int prime = 31;
    // int result = 1;
    // result = prime * result + ((property == null) ? 0 : property.hashCode());
    // result = prime * result + ((reference == null) ? 0 :
    // reference.hashCode());
    // result = prime * result + ((species == null) ? 0 : species.hashCode());
    // return result;
    // }
    // @Override
    // public boolean equals(Object obj) {
    // if (this == obj)
    // return true;
    // if (obj == null)
    // return false;
    // if (getClass() != obj.getClass())
    // return false;
    // SpeciesProperty other = (SpeciesProperty) obj;
    // if (property == null) {
    // if (other.property != null)
    // return false;
    // } else if (!property.equals(other.property))
    // return false;
    // if (reference == null) {
    // if (other.reference != null)
    // return false;
    // } else if (!reference.equals(other.reference))
    // return false;
    // if (species == null) {
    // if (other.species != null)
    // return false;
    // } else if (!species.equals(other.species))
    // return false;
    // return true;
    // }

}
