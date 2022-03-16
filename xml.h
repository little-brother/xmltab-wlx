// Based on https://github.com/markusfisch/libxml
#ifndef _xml_h_
#define _xml_h_

typedef struct xml_attribute {
	/* Argument name */
	char *key;

	/* Argument value */
	char *value;

	/* Pointer to next argument. May be NULL. */
	struct xml_attribute *next;

	/* Pointer to prev argument. May be NULL. */
	struct xml_attribute *prev;
} xml_attribute;

typedef struct xml_element {
	/* The tag name if this is a tag element or NULL if this
	 * element represents character data. */
	char *key;

	/* Character data segment of parent XML element or NULL if
	 * this is a tag element. "key" and "value" are exclusive
	 * because character data is treated as "nameless" child
	 * element. So each xml_element can have either a "key"
	 * or a "value" but not both. */
	char *value;

	/* Parent element. May be NULL. */
	struct xml_element *parent;

	/* First child element. May be NULL. */
	struct xml_element *first_child;

	/* Last child element. May be NULL. */
	struct xml_element *last_child;

	/* Pointer to next sibling. May be NULL. */
	struct xml_element *next;
	
	/* Pointer to prev sibling. May be NULL. */
	struct xml_element *prev;	
	
	/* Added: child count, XML offset and XML length */
	int child_count;
	int offset;
	int length;

	/* First and last attribute. Both may be NULL. */
	xml_attribute *first_attribute, *last_attribute;
	
	int attribute_count;
	void* userdata;
} xml_element;

struct xml_element* xml_parse(const char *);
void xml_free(struct xml_element *);

struct xml_element* xml_find_element(struct xml_element*, const char*);
struct xml_attribute* xml_find_attribute(struct xml_element*, const char*);

char* xml_content(struct xml_element *);
char* xml_path(struct xml_element *);

#endif
