namespace Excel.NAnt.Tasks
{
    using System;
    using global::NAnt.Core;
    using global::NAnt.Core.Attributes;
    using global::NAnt.Core.Types;
    using System.Collections;

    public class Worksheets:CollectionBase
    {
        /// <summary>
        /// The default property for the collection
        /// </summary>
        /// <param name="index">The index of the item to retrieve</param>
        /// <returns>The Worksheet object associated with the index provided.</returns>
        public Worksheet this[int index]
        {
            get { return List[index] as Worksheet; }
            set { List[index] = value; }
        }

        /// <summary>
        /// Adds a value to the collection
        /// </summary>
        /// <param name="value">The value to add to the collection</param>
        /// <returns>The number of the items currently in the collection.</returns>
        public int Add(Worksheet value)
        {
            return List.Add(value);
        }

        /// <summary>
        /// Examines the collection to obtain the index of a given value.
        /// </summary>
        /// <param name="value">The value to locate</param>
        /// <returns>Returns the index of the located item or a -1 if not found.</returns>
        public int IndexOf(Worksheet value)
        {
            return List.IndexOf(value);
        }

        /// <summary>
        /// Insert a value into the collection at a specific index.
        /// </summary>
        /// <param name="index">The index to add the item to the collection.</param>
        /// <param name="value">The item to add to the collection.</param>
        public void Insert(int index, Worksheet value)
        {
            List.Insert(index, value);
        }

        /// <summary>
        /// Removes a specific item from the collection.
        /// </summary>
        /// <param name="value">The value to remove</param>
        public void Remove(Worksheet value)
        {
            List.Remove(value);
        }

        /// <summary>
        /// Checks to see if the value is contained in the collection.
        /// </summary>
        /// <param name="value">The value to identify</param>
        /// <returns>True if the item is found else false.</returns>
        public bool Contains(Worksheet value)
        {
            return List.Contains(value);
        }

        /// <summary>
        /// Valides that the object is of type Worksheet.
        /// </summary>
        /// <param name="value">The value to validate.</param>
        protected override void OnValidate(object value)
        {
            if (value.GetType() != typeof(Worksheet))
            {
                throw new ArgumentException("Value must be of type Worksheet");
            }
        }

    }
}
