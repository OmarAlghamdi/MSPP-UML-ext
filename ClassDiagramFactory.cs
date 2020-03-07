using System;
namespace slide_looper
{

    public class ClassDiagramFactory
    {

        private static ClassDiagramFactory INSTANCE = new ClassDiagramFactory();
        private int objectCout = 0;

        private ClassDiagramFactory()
        {
        }

        public static ClassDiagramFactory getFactory()
        {
            return INSTANCE;
        }

        public void createClass()
        {
            using (ClassForm prompt = new ClassForm())
            {
                DialogResult result = prompt.ShowDialog();
                if (result == DialogResult.OK)
                {
                    drawClass(prompt.name, prompt.fields, prompt.ops);
                }
            }

        }

        public void createInterface()
        {

        }

        private void createEntity()
        {

        }

        public void createAssociation()
        {

        }

        public void createAggregation()
        {

        }

        public void createComposition()
        {

        }

        public void createGeneralization()
        {

        }

        public void createImplementation()
        {

        }

        private void createRelation()
        {

        }
    }
}
